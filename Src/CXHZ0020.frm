VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Begin VB.Form CXHZ0020 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "çﬁóøî≠íçâÊñ ì¸óÕ(∂⁄›¿ﬁå`éÆ)"
   ClientHeight    =   9750
   ClientLeft      =   90
   ClientTop       =   1530
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
   ScaleHeight     =   9750
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
      Left            =   7230
      ScaleHeight     =   645
      ScaleWidth      =   7680
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   840
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
      TabIndex        =   56
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
         BackColor       =   8438015
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
         BackColor       =   8438015
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
         BackColor       =   8438015
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
      Height          =   6315
      Left            =   9780
      ScaleHeight     =   6255
      ScaleWidth      =   5130
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   2985
      Width           =   5190
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   780
         Index           =   1
         Left            =   15
         Top             =   45
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   1376
         ForeColor       =   0
         Caption         =   "çﬁ  óø  ñº"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D B1lb_Zainm 
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
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   780
         Index           =   2
         Left            =   15
         Top             =   825
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1376
         ForeColor       =   0
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
         cPositionX      =   30
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_Seban 
            Height          =   360
            Left            =   60
            Top             =   330
            Width           =   1110
            _ExtentX        =   1958
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
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   780
         Index           =   3
         Left            =   1260
         Top             =   825
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1376
         ForeColor       =   0
         Caption         =   "éËîzãÊï™"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D B1lb_TKbn 
            Height          =   360
            Left            =   120
            Top             =   315
            Width           =   795
            _ExtentX        =   1402
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
         Height          =   780
         Index           =   8
         Left            =   15
         Top             =   2790
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1376
         ForeColor       =   0
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
         cPositionX      =   30
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_NaijiSu 
            Height          =   360
            Index           =   0
            Left            =   75
            Top             =   315
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "###,##0.000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            cPositionX      =   -100
         End
      End
      Begin Cis3D_v60.CIS3D Back_Nonyu 
         Height          =   780
         Index           =   4
         Left            =   2310
         Top             =   825
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   1376
         ForeColor       =   0
         Caption         =   "éËîzè§é–"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D B1lb_Nonyu 
            Height          =   360
            Left            =   150
            Top             =   315
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1185
         Index           =   6
         Left            =   1695
         Top             =   1605
         Width           =   1710
         _ExtentX        =   3016
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
            Left            =   180
            Top             =   360
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "####0.000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Left            =   180
            Top             =   720
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "####0.000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Height          =   780
         Index           =   9
         Left            =   1710
         Top             =   2790
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1376
         ForeColor       =   0
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
         cPositionX      =   30
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_NaijiSu 
            Height          =   360
            Index           =   1
            Left            =   75
            Top             =   315
            Width           =   1535
            _ExtentX        =   2699
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "###,##0.000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            cPositionX      =   -100
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1185
         Index           =   7
         Left            =   3405
         Top             =   1605
         Width           =   1725
         _ExtentX        =   3043
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
            Left            =   210
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
            Left            =   210
            Top             =   720
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
         Left            =   15
         Top             =   1605
         Width           =   1680
         _ExtentX        =   2963
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
            Left            =   165
            Top             =   390
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
         Begin Cis3D_v60.CIS3D B1lb_ChgBin 
            Height          =   360
            Left            =   780
            Top             =   750
            Width           =   390
            _ExtentX        =   688
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
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   330
            Index           =   0
            Left            =   1110
            Top             =   780
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   582
            ForeColor       =   0
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
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   2
            cPositionX      =   30
            cPositionY      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   780
         Index           =   10
         Left            =   3405
         Top             =   2790
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1376
         ForeColor       =   0
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
         cPositionX      =   30
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_NaijiSu 
            Height          =   360
            Index           =   2
            Left            =   75
            Top             =   315
            Width           =   1535
            _ExtentX        =   2699
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "###,##0.000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            cPositionX      =   -100
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1890
         Index           =   14
         Left            =   15
         Top             =   4350
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   3334
         ForeColor       =   0
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
         cAlingnment     =   6
         cPositionY      =   30
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   780
         Index           =   13
         Left            =   3405
         Top             =   3570
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1376
         ForeColor       =   0
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
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   30
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   780
         Index           =   11
         Left            =   15
         Top             =   3570
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1376
         ForeColor       =   0
         Caption         =   "ë≈êÿécêî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D B1lb_UZan 
            Height          =   360
            Left            =   75
            Top             =   315
            Width           =   1535
            _ExtentX        =   2699
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "###,##0.000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            cPositionX      =   -100
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   780
         Index           =   12
         Left            =   1710
         Top             =   3570
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1376
         ForeColor       =   0
         Caption         =   "ë≈ êÿ ì˙"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D B1lb_UYmd 
            Height          =   360
            Left            =   120
            Top             =   315
            Width           =   1425
            _ExtentX        =   2514
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -100
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
      Height          =   930
      Left            =   315
      ScaleHeight     =   870
      ScaleWidth      =   13110
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1800
      Width           =   13170
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   780
         Index           =   2
         Left            =   3720
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
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   360
            Left            =   135
            TabIndex        =   4
            Top             =   315
            Width           =   375
            _ExtentX        =   661
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
            Text            =   "XX"
            MaxLength       =   2
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
         Caption         =   "éÊ à¯ êÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   90
            TabIndex        =   3
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
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
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
         Left            =   4380
         Top             =   60
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "çﬁ    éø"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iH1_Zaist 
            Height          =   360
            Left            =   90
            TabIndex        =   5
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
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   780
         Index           =   4
         Left            =   11430
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
         Begin CisYM_V60.CisYM iH1_Ym 
            Height          =   360
            Left            =   120
            TabIndex        =   9
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
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   780
         Index           =   0
         Left            =   7740
         Top             =   60
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "î¬ å˙"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iH1_Itatu 
            Height          =   360
            Left            =   135
            TabIndex        =   6
            Top             =   315
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            cFormat         =   "###0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "###0.00"
            MaxLength       =   7
            cDataType       =   2
            cDataReplace    =   1
            cFbComma        =   0
            cFcDecimal      =   2
            cFdAutoFormat   =   1
            cGFormat        =   "###0.00"
            cILength        =   4
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   780
         Index           =   5
         Left            =   8970
         Top             =   60
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "ïù"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iH1_Width 
            Height          =   360
            Left            =   135
            TabIndex        =   7
            Top             =   315
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            cFormat         =   "###0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "###0.00"
            MaxLength       =   7
            cDataType       =   2
            cDataReplace    =   1
            cFbComma        =   0
            cFcDecimal      =   2
            cFdAutoFormat   =   1
            cGFormat        =   "###0.00"
            cILength        =   4
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   780
         Index           =   6
         Left            =   10200
         Top             =   60
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "í∑ Ç≥"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iH1_Long 
            Height          =   360
            Left            =   135
            TabIndex        =   8
            Top             =   315
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            cFormat         =   "###0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "###0.00"
            MaxLength       =   7
            cDataType       =   2
            cDataReplace    =   1
            cFbComma        =   0
            cFcDecimal      =   2
            cFdAutoFormat   =   1
            cGFormat        =   "###0.00"
            cILength        =   4
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
      TabIndex        =   1
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
      Height          =   6285
      Left            =   330
      ScaleHeight     =   6225
      ScaleWidth      =   9270
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3000
      Width           =   9330
      Begin Cis3D_v60.CIS3D BLB_Sur 
         Height          =   6135
         Index           =   1
         Left            =   45
         Top             =   45
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   10821
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
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   900
            Index           =   36
            Left            =   1365
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   5025
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   36
               Left            =   0
               TabIndex        =   46
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   36
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   35
            Left            =   75
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   5025
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   35
               Left            =   0
               TabIndex        =   45
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   35
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   34
            Left            =   7815
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   4125
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   34
               Left            =   0
               TabIndex        =   44
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   34
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   33
            Left            =   6525
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   4125
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   33
               Left            =   0
               TabIndex        =   43
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   33
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   32
            Left            =   5235
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   4125
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   32
               Left            =   0
               TabIndex        =   42
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   32
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   31
            Left            =   3945
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   4125
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   31
               Left            =   0
               TabIndex        =   41
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   31
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   30
            Left            =   2655
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   4125
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   30
               Left            =   0
               TabIndex        =   40
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   30
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   29
            Left            =   1365
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   4125
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   29
               Left            =   0
               TabIndex        =   39
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   29
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   28
            Left            =   75
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   4125
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   28
               Left            =   0
               TabIndex        =   38
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   28
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   27
            Left            =   7815
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   3225
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   27
               Left            =   0
               TabIndex        =   37
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   27
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   26
            Left            =   6525
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   3225
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   26
               Left            =   0
               TabIndex        =   36
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   26
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   25
            Left            =   5235
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   3225
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   25
               Left            =   0
               TabIndex        =   35
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   25
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   24
            Left            =   3945
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   3225
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   24
               Left            =   0
               TabIndex        =   34
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   24
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   23
            Left            =   2655
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   3225
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   23
               Left            =   0
               TabIndex        =   33
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   23
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   22
            Left            =   1365
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   3225
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   22
               Left            =   0
               TabIndex        =   32
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   22
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   21
            Left            =   75
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   3225
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   21
               Left            =   0
               TabIndex        =   31
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   21
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   20
            Left            =   7815
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   2325
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   20
               Left            =   0
               TabIndex        =   30
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   20
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   19
            Left            =   6525
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   2325
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   19
               Left            =   0
               TabIndex        =   29
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   19
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   18
            Left            =   5235
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   2325
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   18
               Left            =   0
               TabIndex        =   28
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   18
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   17
            Left            =   3945
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   2325
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   17
               Left            =   0
               TabIndex        =   27
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   17
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   16
            Left            =   2655
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   2325
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   16
               Left            =   0
               TabIndex        =   26
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   16
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   15
            Left            =   1365
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   2325
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   15
               Left            =   0
               TabIndex        =   25
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   15
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   14
            Left            =   75
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   2325
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   14
               Left            =   0
               TabIndex        =   24
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   14
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   13
            Left            =   7815
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   1425
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   13
               Left            =   0
               TabIndex        =   23
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   13
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   12
            Left            =   6525
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   1425
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   12
               Left            =   0
               TabIndex        =   22
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   12
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   11
            Left            =   5235
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   1425
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   11
               Left            =   0
               TabIndex        =   21
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   11
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   10
            Left            =   3945
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   1425
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   10
               Left            =   0
               TabIndex        =   20
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   10
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   9
            Left            =   2655
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   1425
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   9
               Left            =   0
               TabIndex        =   19
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   9
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   8
            Left            =   1365
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   1425
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   8
               Left            =   0
               TabIndex        =   18
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   8
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   7
            Left            =   75
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   1425
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   7
               Left            =   0
               TabIndex        =   17
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   7
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   37
            Left            =   7110
            ScaleHeight     =   900
            ScaleWidth      =   2010
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   5040
            Width           =   2010
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   300
               Index           =   37
               Left            =   -15
               Top             =   -15
               Width           =   2040
               _ExtentX        =   3598
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
               cPositionX      =   60
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_KakuKei 
               Height          =   300
               Left            =   0
               Top             =   585
               Width           =   2010
               _ExtentX        =   3545
               _ExtentY        =   529
               BackColor       =   16777152
               Caption         =   "##,###,###.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Top             =   285
               Width           =   2010
               _ExtentX        =   3545
               _ExtentY        =   529
               BackColor       =   12632319
               Caption         =   "##,###,###.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   6
            Left            =   7815
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   525
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   6
               Left            =   0
               TabIndex        =   16
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   6
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   5
            Left            =   6525
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   525
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   5
               Left            =   0
               TabIndex        =   15
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   5
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   4
            Left            =   5235
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   525
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   4
               Left            =   0
               TabIndex        =   14
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   4
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   3
            Left            =   3945
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   525
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   3
               Left            =   0
               TabIndex        =   13
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   3
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   2
            Left            =   2655
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   525
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   2
               Left            =   0
               TabIndex        =   12
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   2
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   0
            Left            =   75
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   525
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   0
               Left            =   0
               TabIndex        =   10
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
               cAlignment      =   1
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   0
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16761087
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   900
            Index           =   1
            Left            =   1365
            ScaleHeight     =   900
            ScaleWidth      =   1290
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   525
            Width           =   1290
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   1
               Left            =   0
               TabIndex        =   11
               Top             =   285
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               cFormat         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#####0.000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
               cAlignment      =   1
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   1
               Left            =   -15
               Top             =   -15
               Width           =   1305
               _ExtentX        =   2302
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
               Top             =   600
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "#####0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Left            =   75
            Top             =   75
            Width           =   1290
            _ExtentX        =   2275
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
            Left            =   1365
            Top             =   75
            Width           =   1290
            _ExtentX        =   2275
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
            Left            =   2655
            Top             =   75
            Width           =   1290
            _ExtentX        =   2275
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
            Left            =   3945
            Top             =   75
            Width           =   1290
            _ExtentX        =   2275
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
            Left            =   5235
            Top             =   75
            Width           =   1290
            _ExtentX        =   2275
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
            Left            =   6525
            Top             =   75
            Width           =   1290
            _ExtentX        =   2275
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
            Left            =   7815
            Top             =   75
            Width           =   1290
            _ExtentX        =   2275
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
      End
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   714
      BackColor       =   8438015
      ForeColor       =   0
      Caption         =   "Åy î≠íçâÊñ ì¸óÕ(∂⁄›¿ﬁ∞å`éÆ)  Åz"
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
         BackColor       =   8438015
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
         Left            =   7410
         Top             =   45
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
   End
End
Attribute VB_Name = "CXHZ0020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :   çﬁóøî≠íçâÊñ ì¸óÕÅiÉJÉåÉìÉ_Å[å`éÆÅj
'**       ÉtÉHÅ[ÉÄID    :   CXHZ0020
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :   2004/02/21  By CIS
'**       ïœ  çX  ì˙    :   2005/11/17  By CIS ÉIÉvÉVÉáÉìí«â¡Åiâ“ì≠ì˙ÉJÉåÉìÉ_Å[Åj
'**                                             ÉJÉåÉìÉ_Å[ï\é¶ÇÕÅA0=é©é–ÉJÉåÉìÉ_Å[ÅA1:édì¸êÊÉJÉåÉìÉ_Å[
'**       ïœ  çX  ì˙    :   2008/04/11  By CIS îwî‘çÜÇWåÖëŒâû
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl

    Dim FormAct                 As Boolean      ' Ã´∞—±∏√®Ãﬁ Ã◊∏ﬁ
    Dim ProcHB                  As String * 2   ' ì¸óÕàÊêÿë÷
    Dim SyoriKB                 As String       ' èàóùãÊï™(êVãK•èCê≥•çÌèú ìô)
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
    Dim wHsuZHG(31)             As Currency     ' î≠íçêîÅ^çﬁóøî≠íçâÊñ ì¸óÕÉèÅ[ÉN
    Dim wHsuZCT(31)             As Currency     ' î≠íçêîÅ^çﬁóøî≠íçÉeÅ[ÉuÉã
    Dim wNsuZCT(31)             As Currency
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
    Call GetNJProc_YM(3)
    iH1_Ym = gn_HNYM

'ÉVÉXÉeÉÄÉIÉvÉVÉáÉìälìæ
    Call GetSYSOption(33)
    
    FormAct = Not FormAct
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
'+      ÉtÉHÅ[ÉÄ    Ås QueryUnloadÉCÉxÉìÉg Åt                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub
'+---------------------+
'+    èI  óπ  ÉL Å[     +
'+---------------------+
Private Sub PB_END_Click()
'  << èIóπèàóù >>
    Call EndProc
    
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
      iH1_Torcd.SetFocus
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    êV  ãK  ÉL Å[(F1)       +
'+----------------------------+
Private Sub PB_ADD_Click()
    If ProcHB = "H1" Then
        If SyoriKB <> "ADD" Then
            SyoriKB = "ADD"
            Call SyoriName(SyoriKB)
        End If
        iH1_Torcd.SetFocus
    End If
End Sub
'+----------------------------+
'+    èC  ê≥  ÉL Å[(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If ProcHB = "H1" Then
        If SyoriKB <> "MNT" Then
            SyoriKB = "MNT"
            Call SyoriName(SyoriKB)
        End If
        iH1_Torcd.SetFocus
    End If
End Sub
'+----------------------------+
'+    çÌ  èú  ÉL Å[(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If ProcHB = "H1" Then
        If SyoriKB <> "DEL" Then
            SyoriKB = "DEL"
        Else
            SyoriKB = ""
        End If
        Call SyoriName(SyoriKB)
        iH1_Torcd.SetFocus
    End If
End Sub
'+----------------------------+
'+    å„  ëﬁ  ÉL Å[(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    If ProcHB = "H1" Then
        Call CisFrm.MousePT(11)

        Call SqlSelectSet("Back")
        Call DBInput("Back")

       'iH1_Torcd.SetFocus

        Call CisFrm.MousePT(1)
    End If
End Sub
'+----------------------------+
'+    ëO  êi  ÉL Å[(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    If ProcHB = "H1" Then
        Call CisFrm.MousePT(11)

        Call SqlSelectSet("Next")
        Call DBInput("Next")

       'iH1_Torcd.SetFocus

        Call CisFrm.MousePT(1)
    End If
End Sub
'+----------------------------+
'+    åü  çı  ÉL Å[(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_Ukeir: Exit Sub
    If PB_Look.Tag = "iH1_Zaist" Then Call Look_Zaist: Exit Sub
End Sub
'+----------------------------+
'+      éÊà¯êÊåüçı            +
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 8            ' Åyî≠íçÅzéËîzêÊ/å_ñÒêÊåüçı
    RV_Call = "C"

    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'+----------------------------+
'+      éÊà¯êÊéÛì¸åüçı        +
'+----------------------------+
Private Sub Look_Ukeir()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 8            ' Åyî≠íçÅzéËîzêÊ/å_ñÒêÊåüçı
    RV_Torcd = iH1_Torcd
    RV_Kojyo = ""
    RV_Call = "C"

    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
        iH1_Ukeir = RV_Ukeir
    End If
End Sub
'+----------------------------+
'+      çﬁóøéÊà¯êÊåüçı        +
'+----------------------------+
Private Sub Look_Zaist()
    RV_Left = 0
    RV_Top = 0
    RVI_SK = 1
    RVI_Torcd = RTrim(iH1_Torcd)
    RVI_Ukeir = RTrim(iH1_Ukeir)
    RVI_Hinmoku = 1

    CKK0077.Show vbModal
    Unload CKK0077
    Set CKK0077 = Nothing

    If RV_Rtn Then
        iH1_Zaist = RV_Zaist
        iH1_Itatu = RV_Itatu
        iH1_Width = RV_Width
        iH1_Long = RV_Long
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
    H1lb_Tornm = ""
    If TorNmGet(iH1_Torcd, 8) Then
       H1lb_Tornm = TRM_RName
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
'+  iH1_Zaist
'+-----------------------------+
Private Sub iH1_Zaist_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Zaist_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Sur
'+-----------------------------+
Private Sub iB1_Sur_LostFocus(INDEX As Integer)
    Call SurKeisan
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
        Case Else
            SyoriNM = ""
    End Select
    For Each gObj In Me.Controls
        If gObj.Name Like "i*" Then gObj.ForeColor = W_Color
        If gObj.Name Like "*Comb*" Then gObj.ForeColor = W_Color
'       * PF-Key ï\é¶ê›íË *
        If SyoriKB <> "LOC" Then
            If gObj.Name = "PB_ADD" Or _
               gObj.Name = "PB_MNT" Or _
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

    H1Mode = False: H1Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
' ¿ﬁ–∞çÄñ⁄(âÊñ ÇÃÉ`ÉâÉcÉLñhé~)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1": H1Mode = True: H1Color = gPGotSel
       Case "B1": B1Mode = True: B1Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode: H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode: B1_Area1.BackColor = B1Color
    B1_Area2.Enabled = B1Mode: B1_Area2.BackColor = B1Color

    PB_ADD.Visible = H1Mode
    PB_MNT.Visible = H1Mode
    PB_DEL.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode

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
            If .Name Like "i" & wbc_Enm & "_*" Then .Text = "": .Enabled = True
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

'   éÊà¯êÊ
    If Trim(iH1_Torcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    éÊà¯êÊÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            .MB_Button = Error
            .MBOX
        End With
        iH1_Torcd.SetFocus
        Exit Function
    Else
        If Not TorNmGet(iH1_Torcd, 8) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    éÊà¯êÊÅ@ñ¢ìoò^ÅIÅiéÊà¯êÊÉ}ÉXÉ^Åj            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_Torcd.SetFocus
           Exit Function
        End If
    End If
'   éÛì¸
    If RTrim(iH1_Ukeir) <> "" Then
        gSL_Select = ""
        gSL_Select = gSL_Select & "select *"
        gSL_Select = gSL_Select & "  from éÊà¯êÊéÛì¸É}ÉXÉ^"
        gSL_Select = gSL_Select & " where éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   and éÊà¯êÊãÊï™ = 1"
        gSL_Select = gSL_Select & "   and éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
        If Not TUMRead(gSL_Select, 1) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    éÛì¸Å@ñ¢ìoò^ÅIÅiéÊà¯êÊéÛì¸É}ÉXÉ^Åj            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Ukeir.SetFocus
            Exit Function
        End If
    End If
'   çﬁéø
    If Trim(iH1_Zaist) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    çﬁéøÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            .MB_Button = Error
            .MBOX
        End With
        iH1_Zaist.SetFocus
        Exit Function
    Else
        gSL_Select = ""
        gSL_Select = gSL_Select & "select *"
        gSL_Select = gSL_Select & "  from çﬁóøéÊà¯êÊÉ}ÉXÉ^"
        gSL_Select = gSL_Select & " where éÊà¯êÊ = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   and éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & "   and çﬁéø = '" & RTrim(iH1_Zaist) & "'"
        gSL_Select = gSL_Select & "   and î¬å˙ = " & CisFun.Val2(iH1_Itatu) & ""
        gSL_Select = gSL_Select & "   and ïù = " & CisFun.Val2(iH1_Width) & ""
        gSL_Select = gSL_Select & "   and í∑Ç≥ = " & CisFun.Val2(iH1_Long) & ""
        If Not ZTMRead(gSL_Select, 1) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    çﬁéøÅEî¬å˙ÅEïùÅEí∑Ç≥Å@ñ¢ìoò^ÅIÅiçﬁóøéÊà¯êÊÉ}ÉXÉ^Åj            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Zaist.SetFocus
            Exit Function
        End If
        If ZTM.SKãÊï™ <> 1 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    Ç©ÇÒÇŒÇÒïîïiÇÕì¸óÕïsâ¬Ç≈Ç∑ÅI            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Zaist.SetFocus
            Exit Function
        End If
    End If
'   îNåé
    If Trim(iH1_Ym) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    îNåéÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_Ym.SetFocus
        Exit Function
    Else
        If Not iH1_Ym.cChkResult Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    îNåéÇê≥ÇµÇ≠ì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Ym.SetFocus
            Exit Function
        End If
        If iH1_Ym < Format(Now(), "yyyymm") Then
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(2) = "    âﬂãéÇÃì˙ïtÇ≈Ç∑ÅIÅ@ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH            "
            CisFun.MB_Button = OK_CAN
            If Not CisFun.MBOX Then
                iH1_Ym.SetFocus
                Exit Function
            End If
        End If
        If gSYSOp_CLKb = 1 Then
            gStr = RTrim(iH1_Torcd)
        Else
            gStr = ""
        End If
        If Not CalendarGet(iH1_Ym.cYear, iH1_Ym.cMonth, 1, gStr) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    " & RTrim(iH1_Ym.cYear) & "îN" & RTrim(iH1_Ym.cMonth) & "åéÇÃÉJÉåÉìÉ_ñ¢ìoò^ÅI            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Ym.SetFocus
            Exit Function
        End If
        
    End If

    Call SqlSelectSet("INV")
    If DBInput("INV") Then
        If SyoriKB = "ADD" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ç°âÒî≠íçÉfÅ[É^Å@ìoò^çœÅI            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Torcd.SetFocus
            Exit Function
        End If
    Else
        If SyoriKB <> "ADD" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ç°âÒî≠íçÉfÅ[É^Å@ñ¢ìoò^ÅI            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Torcd.SetFocus
            Exit Function
        End If
    End If

    Head1Chk = True
End Function
'****************************
'*      Body1¡™Ø∏èàóù
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False

    Call SurKeisan

'   çﬁóøî≠íçë≈êÿÉfÅ[É^ÅH
    If ZUT_RDSTS = True Then
    '   écêîê›íË
        If ZUT.ë≈êÿéÌï  = 0 Then
            If ZUT.ê›íËécêî < CisFun.Val2(B1lb_SurKei) Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    ë≈êÿécêîÅÉî≠íçêîÅI            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Exit Function
            End If
        End If
    End If
    For wIndex = 0 To 36
        If CisFun.Val2(iB1_Sur(wIndex)) <> 0 Then
        '   â“ìÆì˙ÇÃämîF
            If iB1_Sur(wIndex).Tag = "1" Then
                CisFun.MB_Lines = 4
                CisFun.MB_MSG(2) = "    " & B1lb_DD(wIndex).Caption & "ì˙ÇÕîÒâ“ì≠ì˙Ç≈Ç∑ÅIÅ@ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH            "
                CisFun.MB_Button = OK_CAN
                If Not CisFun.MBOX Then
                    iB1_Sur(wIndex).SetFocus
                    Exit Function
                End If
            End If
        '   çﬁóøî≠íçë≈êÿÉfÅ[É^ÅH
            If ZUT_RDSTS = True Then
            '   ë≈êÿì˙Åiçﬁéø.. or çﬁéø..ÅEédì¸êÊ or édì¸êÊÅj
                If Trim(ZUT.ë≈êÿì˙) <> "" Then
                    wYmd = iH1_Ym & Format(B1lb_DD(wIndex).Tag, "00")
                    If ZUT.ë≈êÿì˙ <= wYmd Then
                        With CisFun
                            .MB_Lines = 4
                            .MB_MSG(2) = "    " & B1lb_DD(wIndex).Caption & "ì˙ÇÕë≈êÿçœÇ≈Ç∑ÅI            "
                            .MB_Title = ""
                            .MB_Button = Error
                            .MBOX
                        End With
                        iB1_Sur(wIndex).SetFocus
                        Exit Function
                    End If
                End If
            End If
        End If
    Next wIndex

    Body1Chk = True
End Function
'**************************************
'*                                    *
'*       ÉfÅ[É^ÉxÅ[ÉXì¸óÕ Åï ï\é¶       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:

    Call HeadBodyClear("B")

    If Not ZHGRead(gSL_Select, 1) Then
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
        iH1_Torcd = RTrim(ZHG.édì¸êÊ)
        iH1_Ukeir = RTrim(ZHG.éÛì¸)
        iH1_Zaist = RTrim(ZHG.çﬁéø)
        iH1_Itatu = RTrim(ZHG.î¬å˙)
        iH1_Width = RTrim(ZHG.ïù)
        iH1_Long = RTrim(ZHG.í∑Ç≥)
        iH1_Ym = CisFun.Mid2(ZHG.î[ì¸ì˙, 1, 6)
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
    gSL_Select = gSL_Select & "  zhg.édì¸êÊ"
    gSL_Select = gSL_Select & ", zhg.éÛì¸"
    gSL_Select = gSL_Select & ", zhg.çﬁéø"
    gSL_Select = gSL_Select & ", zhg.î¬å˙"
    gSL_Select = gSL_Select & ", zhg.ïù"
    gSL_Select = gSL_Select & ", zhg.í∑Ç≥"
    gSL_Select = gSL_Select & ", zhg.î[ì¸ì˙"
    gSL_Select = gSL_Select & ", zhg.î≠íçêî"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  çﬁóøî≠íçâÊñ ì¸óÕÉèÅ[ÉN zhg"
    gSL_Select = gSL_Select & " where zhg.èàóùí[ññ = substring(host_name(), 1, 20)"
    gSL_Select = gSL_Select & "   and zhg.î≠íçãÊï™ = '5'"
    gSL_Select = gSL_Select & "   and"
    gSL_Select = gSL_Select & " ( zhg.édì¸êÊ + space(7 - len(zhg.édì¸êÊ))"
    gSL_Select = gSL_Select & " + zhg.éÛì¸ + space(2 - len(zhg.éÛì¸))"
    gSL_Select = gSL_Select & " + zhg.çﬁéø + space(25 - len(zhg.çﬁéø))"
    gSL_Select = gSL_Select & " + right('0000' + cast(cast(zhg.î¬å˙ as decimal(6,2)) as varchar(7)), 7)"
    gSL_Select = gSL_Select & " + right('0000' + cast(cast(zhg.ïù as decimal(6,2)) as varchar(7)), 7)"
    gSL_Select = gSL_Select & " + right('0000' + cast(cast(zhg.í∑Ç≥ as decimal(6,2)) as varchar(7)), 7)"
    gSL_Select = gSL_Select & " + substring(zhg.î[ì¸ì˙, 1, 6)"
    gSL_Select = gSL_Select & " ) "
    Select Case InputKey
        Case "Back": gSL_Select = gSL_Select & "<"
        Case "Next": gSL_Select = gSL_Select & ">"
        Case Else:   gSL_Select = gSL_Select & "="
    End Select
    gSL_Select = gSL_Select & " ( '"
    gSL_Select = gSL_Select & iH1_Torcd & Space(7 - CisFun.Len2(iH1_Torcd))
    gSL_Select = gSL_Select & iH1_Ukeir & Space(2 - CisFun.Len2(iH1_Ukeir))
    gSL_Select = gSL_Select & iH1_Zaist & Space(25 - CisFun.Len2(iH1_Zaist))
    gSL_Select = gSL_Select & Format(CisFun.Val2(iH1_Itatu), "0000.00")
    gSL_Select = gSL_Select & Format(CisFun.Val2(iH1_Width), "0000.00")
    gSL_Select = gSL_Select & Format(CisFun.Val2(iH1_Long), "0000.00")
    gSL_Select = gSL_Select & iH1_Ym & Space(6 - CisFun.Len2(iH1_Ym))
    gSL_Select = gSL_Select & "' )"
    Select Case InputKey
        Case "Back"
            gSL_Select = gSL_Select & " order by"
            gSL_Select = gSL_Select & "  zhg.édì¸êÊ desc"
            gSL_Select = gSL_Select & ", zhg.éÛì¸ desc"
            gSL_Select = gSL_Select & ", zhg.çﬁéø desc"
            gSL_Select = gSL_Select & ", zhg.î¬å˙ desc"
            gSL_Select = gSL_Select & ", zhg.ïù desc"
            gSL_Select = gSL_Select & ", zhg.í∑Ç≥ desc"
            gSL_Select = gSL_Select & ", zhg.î[ì¸ì˙ desc"
        Case "Next"
            gSL_Select = gSL_Select & " order by"
            gSL_Select = gSL_Select & "  zhg.édì¸êÊ asc"
            gSL_Select = gSL_Select & ", zhg.éÛì¸ asc"
            gSL_Select = gSL_Select & ", zhg.çﬁéø asc"
            gSL_Select = gSL_Select & ", zhg.î¬å˙ desc"
            gSL_Select = gSL_Select & ", zhg.ïù desc"
            gSL_Select = gSL_Select & ", zhg.í∑Ç≥ desc"
            gSL_Select = gSL_Select & ", zhg.î[ì¸ì˙ asc"
    End Select
'
End Sub
'**************************************
'*          BODY ì‡óeæØƒèàóù           *
'**************************************
Private Sub BodySet(SetType As String)
    Dim wYm(2)      As String
'
    With CisFun
        wYm(0) = iH1_Ym
        .Date_RsFormat = fYM
        For gInt = 1 To 2
            .Date_Add iH1_Ym, M, gInt
            wYm(gInt) = .Date_Result
        Next gInt
    End With

    For wIndex = 0 To 31
        wHsuZHG(wIndex) = 0
        wHsuZCT(wIndex) = 0
        wNsuZCT(wIndex) = 0
    Next wIndex

'   çﬁóøî≠íçâÊñ ì¸óÕÉèÅ[ÉN
    Call SqlSelectSet("")
    If ZHGRead(gSL_Select) Then
        Do Until Not ZHG_RDSTS
            wIndex = CisFun.Mid2(ZHG.î[ì¸ì˙, 7, 2)
            wHsuZHG(wIndex) = wHsuZHG(wIndex) + ZHG.î≠íçêî
            wHsuZHG(0) = wHsuZHG(0) + ZHG.î≠íçêî

            Call ZHGReadNext
        Loop
        B1lb_SurKei = Format(wHsuZHG(0), "#,##0.000")
    End If
    Call ZHGClose
'   çﬁóøî≠íçÉeÅ[ÉuÉã
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  zct.î[ì¸ì˙"
    gSL_Select = gSL_Select & ", zct.î≠íçêî"
    gSL_Select = gSL_Select & ", case when zyt.ì¸â◊êî is null then 0 else 1 end åèêî"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  çﬁóøî≠íçÉeÅ[ÉuÉã zct"
    gSL_Select = gSL_Select & "  left outer join çﬁóøì¸â◊ÉeÅ[ÉuÉã zyt"
    gSL_Select = gSL_Select & "    on zyt.î≠íçä«óùNO = zct.î≠íçä«óùNO"
    gSL_Select = gSL_Select & " where zct.édì¸êÊ = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   and zct.éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   and zct.çﬁéø = '" & RTrim(iH1_Zaist) & "'"
    gSL_Select = gSL_Select & "   and zct.î¬å˙ = " & CisFun.Val2(iH1_Itatu) & ""
    gSL_Select = gSL_Select & "   and zct.ïù = " & CisFun.Val2(iH1_Width) & ""
    gSL_Select = gSL_Select & "   and zct.í∑Ç≥ = " & CisFun.Val2(iH1_Long) & ""
    gSL_Select = gSL_Select & "   and substring(zct.î[ì¸ì˙, 1, 6) = '" & iH1_Ym & "'"
    If ZCTRead(gSL_Select) Then
        Do Until Not ZCT_RDSTS
            wIndex = CisFun.Mid2(ZCT.î[ì¸ì˙, 7, 2)
            wNsuZCT(wIndex) = wNsuZCT(wIndex) + ZCT.åèêî
            wHsuZCT(wIndex) = wHsuZCT(wIndex) + ZCT.î≠íçêî
            wHsuZCT(0) = wHsuZCT(0) + ZCT.î≠íçêî

            Call ZCTReadNext
        Loop
        B1lb_KakuKei = Format(wHsuZCT(0), "#,##0.000")
    End If
    Call ZCTClose

'   <><><><><> ì˙ïtèÓïÒæØƒ <><><><><>
    Call DaySet

'   çﬁóøéÊà¯êÊÉ}ÉXÉ^
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  ztm.çﬁóøä«óùî‘çÜ"
    gSL_Select = gSL_Select & ", ztm.îwî‘çÜ"
    gSL_Select = gSL_Select & ", ztm.éËîzãÊï™"
    gSL_Select = gSL_Select & ", ztm.SKãÊï™"
    gSL_Select = gSL_Select & ", ztm.ÉTÉCÉNÉã"
    gSL_Select = gSL_Select & ", ztm.é˚óeêîÇP"
    gSL_Select = gSL_Select & ", ztm.é˚óeäÌÇP"
    gSL_Select = gSL_Select & ", ztm.ïœçXì˙"
    gSL_Select = gSL_Select & ", ztm.ïœçXï÷"
    gSL_Select = gSL_Select & ", ztm.é˚óeêîÇQ"
    gSL_Select = gSL_Select & ", ztm.é˚óeäÌÇQ"
    gSL_Select = gSL_Select & ", ztm.å_ñÒêÊ"
    gSL_Select = gSL_Select & ", isnull(zrm.çﬁóøñº, '') çﬁóøñº"
    gSL_Select = gSL_Select & ", isnull(sym.ílñºèÃ, '') éËîzãÊï™ñº"
    gSL_Select = gSL_Select & ", isnull(trm.ó™èÃ, '') ÉÅÅ[ÉJÅ[ñº"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  çﬁóøéÊà¯êÊÉ}ÉXÉ^ ztm"
    gSL_Select = gSL_Select & "  left outer join çﬁóøÉ}ÉXÉ^ zrm"
    gSL_Select = gSL_Select & "    on zrm.çﬁóøä«óùî‘çÜ = ztm.çﬁóøä«óùî‘çÜ"
    gSL_Select = gSL_Select & "  left outer join éÊà¯êÊÉ}ÉXÉ^ trm"
    gSL_Select = gSL_Select & "    on trm.éÊà¯êÊCD = ztm.ÉÅÅ[ÉJÅ["
    gSL_Select = gSL_Select & "   and trm.éÊà¯êÊãÊï™ = 1"
    gSL_Select = gSL_Select & "  left outer join ñºèÃÉ}ÉXÉ^ sym"
    gSL_Select = gSL_Select & "    on sym.ãÊï™ñºèÃ = 'éËîzãÊï™'"
    gSL_Select = gSL_Select & "   and sym.ãÊï™É^ÉCÉv = ''"
    gSL_Select = gSL_Select & "   and sym.íl = ztm.éËîzãÊï™"
    gSL_Select = gSL_Select & " where ztm.éÊà¯êÊ = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   and ztm.éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   and ztm.çﬁéø = '" & RTrim(iH1_Zaist) & "'"
    gSL_Select = gSL_Select & "   and ztm.î¬å˙ = " & CisFun.Val2(iH1_Itatu) & ""
    gSL_Select = gSL_Select & "   and ztm.ïù = " & CisFun.Val2(iH1_Width) & ""
    gSL_Select = gSL_Select & "   and ztm.í∑Ç≥ = " & CisFun.Val2(iH1_Long) & ""
    If ZTMRead(gSL_Select, 1) Then
        With ZTM
            B1lb_Zainm = RTrim(.çﬁóøñº)
            B1lb_Seban = RTrim(.îwî‘çÜ)
            B1lb_TKbn = RTrim(.éËîzãÊï™ñº)
            B1lb_Nonyu = RTrim(.ÉÅÅ[ÉJÅ[ñº)
            If Trim(.ïœçXì˙) <> "" Then
                B1lb_ChgDay = Format(.ïœçXì˙, "0000/00/00")
                B1lb_ChgBin = RTrim(.ïœçXï÷)
            End If
            If .é˚óeêî1 <> 0 Then B1lb_Syuyo(1) = Format(.é˚óeêî1, "#,##0.000")
            If .é˚óeêî2 <> 0 Then B1lb_Syuyo(2) = Format(.é˚óeêî2, "#,##0.000")
            B1lb_Youki(1) = RTrim(.é˚óeäÌ1)
            B1lb_Youki(2) = RTrim(.é˚óeäÌ2)
        End With
    Else
        Exit Sub    '<><><> çﬁóøéÊà¯êÊÉ}ÉXÉ^Å@ñ¢ìoò^éûÇÕâ∫ãLÇÃèàóùÇÕñ¢é¿çsÅI <><><>
    End If
'   çﬁóøî≠íçì‡é¶ÉeÅ[ÉuÉã
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  sum(case when substring(znj.î[ä˙, 1, 6) = '" & wYm(0) & "' then znj.êîó  else 0 end) ìñåéì‡é¶êî"
    gSL_Select = gSL_Select & ", sum(case when substring(znj.î[ä˙, 1, 6) = '" & wYm(1) & "' then znj.êîó  else 0 end) óÇåéì‡é¶êî"
    gSL_Select = gSL_Select & ", sum(case when substring(znj.î[ä˙, 1, 6) = '" & wYm(2) & "' then znj.êîó  else 0 end) óÇÅXåéì‡é¶êî"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  çﬁóøî≠íçì‡é¶ÉeÅ[ÉuÉã znj"
    gSL_Select = gSL_Select & " where znj.ì‡é¶îNåé = '" & iH1_Ym & "'"
    gSL_Select = gSL_Select & "   and znj.çﬁóøä«óùî‘çÜ = " & ZTM.çﬁóøä«óùî‘çÜ & ""
    gSL_Select = gSL_Select & "   and znj.éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   and znj.éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    If ZNJRead(gSL_Select, 1) Then
        With ZNJ
            B1lb_NaijiSu(0) = Format(.ìñåéì‡é¶êî, "#,###")
            B1lb_NaijiSu(1) = Format(.óÇåéì‡é¶êî, "#,###")
            B1lb_NaijiSu(2) = Format(.óÇÅXåéì‡é¶êî, "#,###")
        End With
    End If
'   çﬁóøî≠íçë≈êÿÉeÅ[ÉuÉã
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  zut.ë≈êÿéÌï "
    gSL_Select = gSL_Select & ", zut.åxçêêî"
    gSL_Select = gSL_Select & ", zut.ê›íËécêî"
    gSL_Select = gSL_Select & ", zut.ë≈êÿì˙"
    gSL_Select = gSL_Select & ", zut.åxçêì˙"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  çﬁóøî≠íçë≈êÿÉeÅ[ÉuÉã zut"
    gSL_Select = gSL_Select & " where ( zut.ë≈êÿéÌï  = 0" _
                                & " and zut.çﬁóøä«óùî‘çÜ = " & ZTM.çﬁóøä«óùî‘çÜ _
                                & " )"
    gSL_Select = gSL_Select & "    or ( zut.ë≈êÿéÌï  = 1" _
                                & " and zut.çﬁóøä«óùî‘çÜ = " & ZTM.çﬁóøä«óùî‘çÜ _
                                & " )"
    gSL_Select = gSL_Select & "    or ( zut.ë≈êÿéÌï  = 2" _
                                & " and zut.çﬁóøä«óùî‘çÜ = " & ZTM.çﬁóøä«óùî‘çÜ _
                                & " and zut.édì¸êÊ = '" & RTrim(iH1_Torcd) & "'" _
                                & " )"
    gSL_Select = gSL_Select & "    or ( zut.ë≈êÿéÌï  = 3" _
                                & " and zut.édì¸êÊ = '" & RTrim(iH1_Torcd) & "'" _
                                & " )"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  zut.ë≈êÿéÌï "
    If ZUTRead(gSL_Select, 1) Then
        With ZUT
            If .ë≈êÿéÌï  = 0 Then
                B1lb_UZan = Format(.ê›íËécêî, "#,##0.000")
            End If
            If Trim(.ë≈êÿì˙) <> "" Then
                B1lb_UYmd = Format(.ë≈êÿì˙, "0000/00/00")
            End If
        End With
        ZUT_RDSTS = True
    Else
        ZUT_RDSTS = False
    End If

End Sub
'+-------------------------+
'+  ì˙ïtèÓïÒÉZÉbÉg
'+-------------------------+
Private Sub DaySet()
    Dim wSetDay     As Long
    Dim wSetSta     As Long
    Dim wDate       As Date

    If Trim(iH1_Ym) = "" Then Exit Sub

'   *  ÉJÉåÉìÉ_ïîì˙ït ÉZÉbÉg  *
    If gSYSOp_CLKb = 1 Then
        gStr = RTrim(iH1_Torcd)
    Else
        gStr = ""
    End If
    If Not CalendarGet(iH1_Ym.cYear, iH1_Ym.cMonth, 1, gStr) Then
        Call ItemsClearCLM
        CLM.â“ìÆãÊï™ = String(31, "0")
        CLM.ñÈãŒãÊï™ = String(31, "0")
        With CisFun
            .Date_RsFormat = fDD
            .Date_End iH1_Ym
            CLM.ññì˙ = .Val2(.Date_Result)
            CLM.â“ìÆì˙êî = CLM.ññì˙
        End With
    End If

    wDate = DateValue(iH1_Ym.cYear & "/" & iH1_Ym.cMonth & "/01")
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
                wDate = DateValue(iH1_Ym.cYear & "/" & iH1_Ym.cMonth & "/" & wSetDay)
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
            If wHsuZHG(wSetDay) <> 0 Then
                iB1_Sur(wIndex) = wHsuZHG(wSetDay)
            End If
            If wHsuZCT(wSetDay) <> 0 Then
                B1lb_Kaku(wIndex) = Format(wHsuZCT(wSetDay), "#####0.000")
            End If
'           ì¸â◊óLÇÃéûÇÕîwåiêFÇïœçX
            If wNsuZCT(wSetDay) <> 0 Then
                B1lb_Kaku(wIndex).BackColor = mColorNYT
            End If
            B_Sur(wIndex).Visible = True
        End If
    Next wIndex

End Sub
'+-------------------------+
'+  êîó åvéZÅïï\é¶
'+-------------------------+
Private Sub SurKeisan()
    wHsuZHG(0) = 0
    For wIndex = 0 To 36
        wHsuZHG(0) = wHsuZHG(0) + CisFun.Val2(iB1_Sur(wIndex))
    Next wIndex
    If wHsuZHG(0) <> 0 Then
        B1lb_SurKei = Format(wHsuZHG(0), "#,##0.000")
    Else
        B1lb_SurKei = ""
    End If
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
    Call DBDelete_ZHG
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
        If SyoriKB = "ADD" Then
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
    Call DBDelete_ZHG
' ( í«â¡èàóù )
    Call DBSet_ZHG
    For wIndex = 0 To 36
        If CisFun.Val2(iB1_Sur(wIndex)) <> 0 Then
            Call DBInsert_ZHG(wIndex)
        End If
    Next wIndex
'   åèêîï\é¶
    Select Case SyoriKB
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
'+-----------------------------------------+
'+      çÌèúèàóù - çﬁóøî≠íçâÊñ ì¸óÕÉèÅ[ÉN
'+-----------------------------------------+
Private Function DBDelete_ZHG(Optional DBNo As Integer) As Boolean
    DBDelete_ZHG = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "delete çﬁóøî≠íçâÊñ ì¸óÕÉèÅ[ÉN"
    gSL_Select = gSL_Select & " where èàóùí[ññ = substring(host_name(), 1, 20)"
    gSL_Select = gSL_Select & "   and î≠íçãÊï™ = '5'"
    gSL_Select = gSL_Select & "   and édì¸êÊ = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   and éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   and substring(î[ì¸ì˙, 1, 6) = '" & iH1_Ym & "'"
    gSL_Select = gSL_Select & "   and çﬁéø = '" & RTrim(iH1_Zaist) & "'"
    gSL_Select = gSL_Select & "   and î¬å˙ = " & CisFun.Val2(iH1_Itatu) & ""
    gSL_Select = gSL_Select & "   and ïù = " & CisFun.Val2(iH1_Width) & ""
    gSL_Select = gSL_Select & "   and í∑Ç≥ = " & CisFun.Val2(iH1_Long) & ""
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With

    DBDelete_ZHG = True
End Function
'+-----------------------------------------+
'+      ÉZÉbÉg - çﬁóøî≠íçâÊñ ì¸óÕÉèÅ[ÉN
'+-----------------------------------------+
Private Function DBSet_ZHG()
    DBSet_ZHG = False

    Call ItemsClearZHG
    With ZHG
        .î≠íçãÊï™ = 5
        .édì¸êÊ = RTrim(iH1_Torcd)
        .éÛì¸ = RTrim(iH1_Ukeir)
        .î≠íçì˙ = Format(Date, "yyyymmdd")
        .î≠íçï÷ = ""
        .çﬁóøä«óùî‘çÜ = ZTM.çﬁóøä«óùî‘çÜ
        .çﬁéø = RTrim(iH1_Zaist)
        .î¬å˙ = CisFun.Val2(iH1_Itatu)
        .ïù = CisFun.Val2(iH1_Width)
        .í∑Ç≥ = CisFun.Val2(iH1_Long)
        .îwî‘çÜ = ZTM.îwî‘çÜ
        .éËîzãÊï™ = ZTM.éËîzãÊï™
        .SKãÊï™ = ZTM.SKãÊï™
        .ÉTÉCÉNÉã = ZTM.ÉTÉCÉNÉã
        .å_ñÒêÊ = ZTM.å_ñÒêÊ
        .çÏê¨éÌï  = "C"
        .èàóùãÊï™ = 0
        .çÏê¨é“ = gTanto
    End With

    DBSet_ZHG = True
End Function
'+-----------------------------------------+
'+      í«â¡èàóù - çﬁóøî≠íçâÊñ ì¸óÕÉèÅ[ÉN
'+-----------------------------------------+
Private Function DBInsert_ZHG(ID As Long, Optional DBNo As Integer) As Boolean
    DBInsert_ZHG = False

    With ZHG
        .î[ì¸ì˙ = iH1_Ym & Format(B1lb_DD(wIndex).Tag, "00")
        .î[ì¸ï÷ = ""
        If Trim(ZTM.ïœçXì˙) = "" Or _
           Trim(ZTM.ïœçXì˙) > .î[ì¸ì˙ Then
            .é˚óeêî = ZTM.é˚óeêî1
        Else
            .é˚óeêî = ZTM.é˚óeêî2
        End If
        .î≠íçêî = CisFun.Val2(iB1_Sur(ID))
        If .é˚óeêî = 0 Then
            .ñáêî = 0
            .í[êî = 0
        Else
           '.ñáêî = .î≠íçêî \ .é˚óeêî
           '.í[êî = .î≠íçêî Mod .é˚óeêî
           'If .í[êî <> 0 Then .ñáêî = .ñáêî + 1
            .ñáêî = Round(.î≠íçêî / .é˚óeêî)
            .í[êî = .î≠íçêî - (.é˚óeêî * .ñáêî)
            If .í[êî > 0 Then .ñáêî = .ñáêî + 1
            If .í[êî < 0 Then .í[êî = .í[êî + .é˚óeêî
        End If
'       ## ë≈êÿèÓïÒÇÃæØƒ ##
        If ZUT_RDSTS = True Then
            .í[êîãÊï™ = 0
            .ë≈êÿéÌï  = ZUT.ë≈êÿéÌï 
            Select Case .ë≈êÿéÌï 
                Case 0
                '   >>> ê›íËécêîÇÃåvéZ
                    ZUT.ê›íËécêî = ZUT.ê›íËécêî - .î≠íçêî
                    If ZUT.ê›íËécêî = 0 Then
                        .ë≈êÿãÊï™ = 9
                        .ë≈êÿécêî = 0
                    Else
                        If ZUT.åxçêêî >= ZUT.ê›íËécêî Then
                            .ë≈êÿãÊï™ = 1
                            .ë≈êÿécêî = ZUT.ê›íËécêî
                        Else
                            .ë≈êÿãÊï™ = 0
                            .ë≈êÿécêî = 0
                        End If
                    End If
                    .ë≈êÿì˙ = ""
                Case Else
                '   >>> ë≈êÿì˙
                    If ZUT.åxçêì˙ <= .î[ì¸ì˙ Then
                        .ë≈êÿãÊï™ = 1
                        .ë≈êÿì˙ = ZUT.ë≈êÿì˙
                    Else
                        .ë≈êÿãÊï™ = 0
                        .ë≈êÿì˙ = ""
                    End If
                    .ë≈êÿécêî = 0
            End Select
        End If
    End With
    Call ZHGInsert

    DBInsert_ZHG = True
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
    gSL_Select = gSL_Select & "   and clm.éÊà¯êÊ in ('" & RTrim(Torcd) & "', '')"
    gSL_Select = gSL_Select & "   and clm.îN = '" & DateY & "'"
    gSL_Select = gSL_Select & "   and clm.åé = '" & DateM & "'"
    gSL_Select = gSL_Select & "   and clm.éÛì¸ in ('" & RTrim(Ukeir) & "', '')"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  clm.éÊà¯êÊ desc"
    gSL_Select = gSL_Select & ", clm.éÛì¸ desc"
    If Not CLMRead(gSL_Select, 1) Then Exit Function

    CalendarGet = True
End Function
'==========================================================='
'   èIóπèàóùÅiì¸óÕÉfÅ[É^Ç™Ç†ÇÍÇŒå„ï˚èàóùÇçsÇ§Åj
'==========================================================='
Private Sub EndProc()
    Dim PrtCnt3     As Long
    Dim PrtCnt4     As Long
    Dim OpUchiPrt   As Byte
'ì¸óÕåèêîämîF
    gSL_Select = "SELECT COUNT(*) åèêî FROM çﬁóøî≠íçâÊñ ì¸óÕÉèÅ[ÉN"
    gSL_Select = gSL_Select & " WHERE èàóùí[ññ = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND çÏê¨éÌï  = 'C'"
    If Not ZHGRead(gSL_Select, 1) Then Exit Sub
    If ZHG.åèêî = 0 Then Exit Sub

'çﬁóøî≠íçë≈êÿçXêV
    With CisDB
        .SQL = "çﬁóøî≠íçâÊñ ì¸óÕå„ï˚èàóù"
        .StoadoCount = 5
        .ParaNo = 0: .ParaIO = Returnóp
        .ParaNo = 1: .ParaIO = OutPutóp            'çﬁóøî≠íçë≈êÿÉeÅ[ÉuÉãçXêVåèêî
        .ParaNo = 2: .ParaIO = OutPutóp            'çﬁóøî≠íçë≈êÿåxçêÉäÉXÉgåèêî
        .ParaNo = 3: .ParaIO = OutPutóp            'çﬁóøî≠íçë≈êÿåxçêÉäÉXÉgåèêî
        .ParaNo = 4: .ParaIO = OutPutóp            'ÉGÉâÅ[ÉRÅ[Éh
        .ParaNo = 5: .ParaIO = OutPutóp          'ÉGÉâÅ[ÉRÅ[Éh
'
        If Not .DBStored Then
            MsgBox "ÉXÉgÉAÉhÉGÉâÅ[ "
            GoTo EndProc_Err
        Else
            .ParaNo = 5
            If .ParaValue <> 0 Then GoTo EndProc_Err
            .ParaNo = 3
            PrtCnt3 = .ParaValue
            .ParaNo = 4
            PrtCnt4 = .ParaValue
        End If
    End With
'+---------------------------------------------------------------------------------  2005/11/07 Insert
    ' î≠íçë≈êÿÉäÉXÉgÅ@î≠çs
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 36
    Call IRNRead("", 1)
    OpUchiPrt = CisFun.Val2(Mid(IRN.oyian, 1, 1))
'+---------------------------------------------------------------------------------  2005/11/07 Insert End
'çﬁóøî≠íçë≈êÿåxçêÉäÉXÉgàÛç¸
    If PrtCnt3 <> 0 And OpUchiPrt = 0 Then                                               '2005/11/07í«â¡
        CHHZ0930_CallType = ""
        If Not CKHZ0930LPrint Then Exit Sub
        DoEvents
    End If
'çﬁóøî≠íçë≈êÿÉäÉXÉgàÛç¸
    If PrtCnt4 <> 0 And OpUchiPrt = 0 Then                                               '2005/11/07í«â¡
        CHHZ0940_CallType = ""
        If Not CKHZ0940LPrint Then Exit Sub
        DoEvents
    End If
' çﬁóøî≠íçèàóù
    CisDB.PassLink
    With CisFun
        .ShellApl = gIniExe & "PXHZ0110.exe"
        .ShellAplPara = "G"
        .ShellTask Freeze
    End With


    Exit Sub
EndProc_Err:
    Select Case CisDB.ParaValue
        Case 1: gStr = "    ÅyÅ@çﬁóøÇ©ÇÒÇŒÇÒì«éÊÉGÉâÅ[ÉeÅ[ÉuÉãÅ@çÌèúÉGÉâÅ[ÅIÅz"
        Case 2: gStr = "    ÅyÅ@í†ï[èoóÕÉGÉâÅ[ÉeÅ[ÉuÉãÅ@çÌèúÉGÉâÅ[ÅIÅz"
        Case 3: gStr = "    ÅyÅ@çﬁóøî≠íçë≈êÿÉeÅ[ÉuÉã(äÆóπ)Å@çXêVÉGÉâÅ[ÅIÅz"
        Case 4: gStr = "    ÅyÅ@çﬁóøî≠íçâÊñ ì¸óÕÉèÅ[ÉN(í[êî)Å@çXêVÉGÉâÅ[ÅIÅz"
        Case 5: gStr = "    ÅyÅ@çﬁóøî≠íçë≈êÿÉeÅ[ÉuÉã(ñ¢äÆ)Å@çXêVÉGÉâÅ[ÅIÅz"
        Case 6: gStr = "    ÅyÅ@í†ï[èoóÕÉeÅ[ÉuÉã(åxçê)Å@èoóÕÉGÉâÅ[ÅIÅz"
        Case 7: gStr = "    ÅyÅ@í†ï[èoóÕÉeÅ[ÉuÉã(ë≈êÿ)Å@èoóÕÉGÉâÅ[ÅIÅz"
        Case 8: gStr = "    ÅyÅ@çﬁóøÇ©ÇÒÇŒÇÒì«éÊÉGÉâÅ[ÉeÅ[ÉuÉãÅ@èoóÕÉGÉâÅ[ÅIÅz"
        Case 9: gStr = "    ÅyÅ@çﬁóøî≠íçâÊñ ì¸óÕÉèÅ[ÉN(èàóù)Å@çXêVÉGÉâÅ[ÅIÅz"
        Case 60: gStr = "    ÅyÅ@ÉGÉâÅ[ä«óùî‘çÜÅ@älìæÉGÉâÅ[ÅIÅz"
    End Select
    MsgBox gStr & "  Code = " & CisDB.ParaValue & "     "

End Sub

