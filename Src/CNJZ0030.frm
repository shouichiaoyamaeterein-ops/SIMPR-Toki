VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{8F4FF474-3E18-11D4-AEC2-00000EA57F9E}#1.0#0"; "CISMSG~1.OCX"
Begin VB.Form CNJZ0030 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "çﬁóøî≠íçì‡é¶ì¸óÕ"
   ClientHeight    =   11235
   ClientLeft      =   780
   ClientTop       =   1020
   ClientWidth     =   15360
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ÇlÇr ÉSÉVÉbÉN"
      Size            =   12
      Charset         =   128
      Weight          =   700
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
   ScaleHeight     =   11235
   ScaleMode       =   0  '’∞ªﬁ∞
   ScaleWidth      =   15390
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   6840
      ScaleHeight     =   675
      ScaleWidth      =   8340
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   450
      Width           =   8400
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   5820
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
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
         Left            =   6645
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
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
         Left            =   7470
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
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
            Size            =   12
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
            Size            =   12
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
            Size            =   12
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
         Left            =   3345
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
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
         Left            =   4170
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
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
         Left            =   4995
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
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
      Begin CisBtn_60.CisBtn PB_REP 
         Height          =   585
         Left            =   2520
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnCaption      =   "ïœçXëO"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   16711935
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
         PFBackColor     =   16744703
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
      Height          =   1845
      Left            =   720
      ScaleHeight     =   1785
      ScaleWidth      =   13890
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   9075
      Width           =   13950
      Begin Cis3D_v60.CIS3D CIS3D31 
         Height          =   1650
         Left            =   45
         Top             =   60
         Width           =   13785
         _ExtentX        =   24315
         _ExtentY        =   2910
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderStyle     =   1
         Begin Cis3D_v60.CIS3D CIS3D14 
            Height          =   795
            Left            =   30
            Top             =   825
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "ì‡é¶î≠íç"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   45
            Begin Cis3D_v60.CIS3D B1lb_NaHaKB 
               Height          =   330
               Left            =   -375
               Top             =   360
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   582
               BackColor       =   12648384
               Caption         =   "ÇµÇ»Ç¢"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cAlingnment     =   7
               cBoderStyle     =   1
            End
            Begin CisMsg_v60.CISMSG NaiHatu_Back 
               Height          =   405
               Index           =   2
               Left            =   75
               TabIndex        =   76
               TabStop         =   0   'False
               Top             =   345
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   714
               BackColor       =   16777152
               ForeColor       =   12582912
               Caption         =   "Ç∑ÇÈ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
               cBoderWidth     =   3
               cAlingnment     =   7
               cBoderStyle     =   1
               gGC1_Red        =   255
               gGC2_Green      =   255
               gGC3_Bule       =   0
               gGM1_Red        =   5
               gGM2_Green      =   5
               gGradation      =   -1  'True
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D63 
            Height          =   795
            Left            =   9795
            Top             =   825
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "ëOåéì‡é¶"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   45
            Begin Cis3D_v60.CIS3D B1lb_ZNJSuT 
               Height          =   360
               Left            =   90
               Top             =   360
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   635
               BackColor       =   12648384
               Caption         =   "99,999,999"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D CIS3D60 
            Height          =   795
            Left            =   8985
            Top             =   30
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "åWêî3"
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
            cAlingnment     =   6
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D B1lb_Kesu3 
               Height          =   330
               Left            =   45
               Top             =   390
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   582
               BackColor       =   12648384
               Caption         =   "??"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D58 
            Height          =   795
            Left            =   8220
            Top             =   30
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "åWêî2"
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
            cAlingnment     =   6
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D B1lb_Kesu2 
               Height          =   330
               Left            =   45
               Top             =   390
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   582
               BackColor       =   12648384
               Caption         =   "??"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D52 
            Height          =   795
            Left            =   11775
            Top             =   825
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "ì‡é¶ç∑"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   45
            Begin Cis3D_v60.CIS3D B1lb_SaNJSuT 
               Height          =   360
               Left            =   90
               Top             =   360
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   635
               BackColor       =   12648384
               Caption         =   "99,999,999"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D CIS3D50 
            Height          =   795
            Left            =   4485
            Top             =   825
            Width           =   3330
            _ExtentX        =   5874
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "óLå¯ä˙å¿"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   45
            Begin Cis3D_v60.CIS3D B1lb_SDate 
               Height          =   360
               Left            =   90
               Top             =   360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   635
               BackColor       =   12648384
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cBoderStyle     =   1
            End
            Begin Cis3D_v60.CIS3D B1lb_EDate 
               Height          =   360
               Left            =   1875
               Top             =   360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   635
               BackColor       =   12648384
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cBoderStyle     =   1
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   300
               Index           =   53
               Left            =   1530
               Top             =   420
               Width           =   270
               _ExtentX        =   476
               _ExtentY        =   529
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
         End
         Begin Cis3D_v60.CIS3D CIS3D48 
            Height          =   795
            Left            =   9750
            Top             =   30
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "ëOåéññç›å…"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D B1lb_Zaiko 
               Height          =   330
               Left            =   45
               Top             =   390
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               BackColor       =   12648384
               Caption         =   "9,999,999"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D CIS3D34 
            Height          =   795
            Left            =   7455
            Top             =   30
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "åWêî1"
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
            cAlingnment     =   6
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D B1lb_Kesu1 
               Height          =   330
               Left            =   45
               Top             =   390
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   582
               BackColor       =   12648384
               Caption         =   "??"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D13 
            Height          =   795
            Left            =   3645
            Top             =   825
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "î‰ó¶"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   45
            Begin Cis3D_v60.CIS3D B1lb_Ritu 
               Height          =   360
               Left            =   75
               Top             =   360
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   635
               BackColor       =   12648384
               Caption         =   "999%"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D38 
            Height          =   795
            Left            =   7815
            Top             =   825
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "ìñåéì‡é¶"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   45
            Begin Cis3D_v60.CIS3D B1lb_NJSuT 
               Height          =   360
               Left            =   90
               Top             =   360
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   635
               BackColor       =   12648384
               Caption         =   "##,###,##0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D CIS3D29 
            Height          =   795
            Left            =   1380
            Top             =   30
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   1402
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D B1lb_Cycle 
               Height          =   330
               Left            =   75
               Top             =   390
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   582
               BackColor       =   12648384
               Caption         =   "9-99-99"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cAlingnment     =   7
               cBoderStyle     =   1
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D27 
            Height          =   795
            Left            =   2550
            Top             =   30
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "é˚óeêîá@"
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
            cAlingnment     =   6
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D B1lb_Syuyo1 
               Height          =   330
               Left            =   60
               Top             =   390
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   582
               BackColor       =   12648384
               Caption         =   "999,999.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D20 
            Height          =   795
            Left            =   4095
            Top             =   30
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "êÿë÷ì˙"
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
            cAlingnment     =   6
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D B1lb_CDate 
               Height          =   330
               Left            =   45
               Top             =   390
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   582
               BackColor       =   12648384
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cBoderStyle     =   1
            End
            Begin Cis3D_v60.CIS3D B1lb_CBin 
               Height          =   330
               Left            =   1350
               Top             =   390
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   582
               BackColor       =   12648384
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   795
            Left            =   11445
            Top             =   30
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "ç›å…î≠íç"
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
            cAlingnment     =   6
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D B1lb_ZKKbn 
               Height          =   330
               Left            =   75
               Top             =   390
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   582
               BackColor       =   12648384
               Caption         =   "î≠íçì_"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Begin Cis3D_v60.CIS3D B1lb_ZKSu 
               Height          =   330
               Left            =   1095
               Top             =   390
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   582
               BackColor       =   12648384
               Caption         =   "999,999"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         End
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   795
            Left            =   5910
            Top             =   30
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "é˚óeêîáA"
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
            cAlingnment     =   6
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D B1lb_Syuyo2 
               Height          =   330
               Left            =   45
               Top             =   390
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   582
               BackColor       =   12648384
               Caption         =   "999,999.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   795
            Left            =   30
            Top             =   30
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "ÇrÇjãÊï™"
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
            cAlingnment     =   6
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D B1lb_SKKbn 
               Height          =   330
               Left            =   -285
               Top             =   390
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   582
               BackColor       =   12648384
               Caption         =   "Ç©ÇÒÇŒÇÒ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cAlingnment     =   7
               cBoderStyle     =   1
            End
            Begin CisMsg_v60.CISMSG NaiHatu_Back 
               Height          =   405
               Index           =   1
               Left            =   75
               TabIndex        =   79
               TabStop         =   0   'False
               Top             =   345
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   714
               BackColor       =   16777152
               ForeColor       =   12582912
               Caption         =   "Ç©ÇÒÇŒÇÒ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
               cBoderWidth     =   3
               cAlingnment     =   7
               cBoderStyle     =   1
               gGC1_Red        =   255
               gGC2_Green      =   255
               gGC3_Bule       =   0
               gGM1_Red        =   5
               gGM2_Green      =   5
               gGradation      =   -1  'True
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   795
            Left            =   1380
            Top             =   825
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   1402
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            Begin CisMsg_v60.CISMSG B1lb_HCSu 
               Height          =   390
               Left            =   60
               TabIndex        =   77
               TabStop         =   0   'False
               Top             =   360
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   688
               BackColor       =   -2147483633
               ForeColor       =   12582912
               Caption         =   "99,999,999.999"
               BackStyle       =   0
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
               cFont3DColor2   =   8421504
               cBoderWidth     =   3
               cAlingnment     =   4
               cBoderStyle     =   1
               gGC1_Red        =   255
               gGC2_Green      =   255
               gGC3_Bule       =   0
               gGM1_Red        =   3
               gGM2_Green      =   3
               gGradation      =   -1  'True
            End
            Begin CisMsg_v60.CISMSG NaiHatu_Back 
               Height          =   315
               Index           =   0
               Left            =   60
               TabIndex        =   78
               TabStop         =   0   'False
               Top             =   60
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               BackColor       =   -2147483633
               ForeColor       =   12582912
               Caption         =   "åªî≠íççœêî"
               BackStyle       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
                  Size            =   11.25
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
               cBoderWidth     =   3
               cAlingnment     =   7
               cBoderStyle     =   1
               gGC1_Red        =   255
               gGC2_Green      =   255
               gGC3_Bule       =   0
               gGM1_Red        =   3
               gGM2_Green      =   3
               gGradation      =   -1  'True
            End
         End
      End
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   1305
      Left            =   713
      ScaleHeight     =   1245
      ScaleWidth      =   13485
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1290
      Width           =   13545
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   285
         Left            =   1560
         Top             =   60
         Width           =   11850
         _ExtentX        =   20902
         _ExtentY        =   503
         ForeColor       =   128
         Caption         =   "ëŒè€îNåé:Ctrl+PageDown/ëOêi Ctrl+PageUp/å„ëﬁ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   1
         cPositionX      =   120
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   825
         Left            =   1560
         Top             =   345
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   1455
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
         cPositionY      =   45
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   375
            Left            =   60
            TabIndex        =   1
            Top             =   390
            Width           =   1350
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
         Height          =   825
         Left            =   9405
         Top             =   345
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   1455
         ForeColor       =   16711680
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
         cPositionY      =   45
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   60
            TabIndex        =   6
            Top             =   375
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
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1050
            Top             =   375
            Width           =   2490
            _ExtentX        =   4392
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
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   300
            Index           =   47
            Left            =   3390
            Top             =   45
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   529
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
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   2
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   375
            Left            =   3570
            TabIndex        =   7
            Top             =   375
            Width           =   360
            _ExtentX        =   635
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
         Height          =   1110
         Left            =   60
         Top             =   60
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   1958
         ForeColor       =   16711680
         Caption         =   "ì‡é¶îNåé"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         cPositionX      =   45
         cPositionY      =   -210
         Begin Cis3D_v60.CIS3D CIS3D12 
            Height          =   585
            Left            =   45
            Top             =   60
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   1032
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "ì‡é¶îNåé"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
            cAlingnment     =   7
            cBoderStyle     =   1
            cPositionX      =   45
         End
         Begin CisYM_V60.CisYM iH1_NJYM 
            Height          =   375
            Left            =   45
            TabIndex        =   8
            Top             =   675
            Width           =   1350
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
      Begin Cis3D_v60.CIS3D BLB_Zaist 
         Height          =   825
         Left            =   3000
         Top             =   345
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   1455
         ForeColor       =   16711680
         Caption         =   "çﬁéø"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         cPositionY      =   45
         Begin CisText_V60.CisText iH1_Zaist 
            Height          =   360
            Left            =   75
            TabIndex        =   2
            Top             =   390
            Width           =   3210
            _ExtentX        =   5662
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
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D19 
         Height          =   825
         Left            =   8370
         Top             =   345
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1455
         ForeColor       =   16711680
         Caption         =   "í∑Ç≥"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         cPositionY      =   45
         Begin CisText_V60.CisText iH1_Long 
            Height          =   360
            Left            =   45
            TabIndex        =   5
            Top             =   390
            Width           =   960
            _ExtentX        =   1693
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
      Begin Cis3D_v60.CIS3D CIS3D17 
         Height          =   825
         Left            =   6300
         Top             =   345
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1455
         ForeColor       =   16711680
         Caption         =   "î¬å˙"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         cPositionY      =   45
         Begin CisText_V60.CisText iH1_Itatu 
            Height          =   360
            Left            =   45
            TabIndex        =   3
            Top             =   390
            Width           =   960
            _ExtentX        =   1693
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
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   825
         Left            =   7335
         Top             =   345
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1455
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
         cPositionY      =   45
         Begin CisText_V60.CisText iH1_Width 
            Height          =   360
            Left            =   45
            TabIndex        =   4
            Top             =   390
            Width           =   960
            _ExtentX        =   1693
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
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'µÃå≈íË
      Left            =   30
      TabIndex        =   0
      Text            =   "Dummy"
      Top             =   390
      Width           =   180
   End
   Begin VB.Timer Timer1 
      Left            =   30
      Top             =   675
   End
   Begin VB.Timer Timer2 
      Left            =   15
      Top             =   1155
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   661
      BackColor       =   8438015
      ForeColor       =   0
      Caption         =   "Åy çﬁóøî≠íçì‡é¶ì¸óÕ  Åz"
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
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   10650
         ScaleHeight     =   345
         ScaleWidth      =   3675
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   0
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
            cPositionX      =   30
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
            cPositionX      =   30
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
            cPositionX      =   30
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
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   345
         Left            =   0
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
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
         Left            =   14400
         Top             =   60
         Width           =   885
         _ExtentX        =   1561
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
         Height          =   285
         Left            =   8100
         Top             =   30
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   503
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
   Begin VB.PictureBox B1_Area1 
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
      Height          =   6270
      Left            =   713
      ScaleHeight     =   6210
      ScaleWidth      =   13875
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2715
      Width           =   13935
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   6075
         Left            =   60
         Top             =   60
         Width           =   13725
         _ExtentX        =   24209
         _ExtentY        =   10716
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
         cBoderStyle     =   1
         Begin Cis3D_v60.CIS3D BLB_Sur 
            Height          =   6015
            Index           =   1
            Left            =   30
            Top             =   30
            Width           =   13680
            _ExtentX        =   24130
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   31
               Left            =   1080
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   74
               TabStop         =   0   'False
               Top             =   4515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   31
                  Left            =   -15
                  TabIndex        =   39
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   31
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇRÇPì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   31
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   31
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   31
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   31
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   30
               Left            =   12420
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   30
                  Left            =   -15
                  TabIndex        =   38
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   30
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇRÇOì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   30
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   30
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   30
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   30
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   29
               Left            =   11160
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   72
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   29
                  Left            =   -15
                  TabIndex        =   37
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   29
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇXì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   29
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   29
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   29
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   29
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   28
               Left            =   9900
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   28
                  Left            =   -15
                  TabIndex        =   36
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   28
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇWì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   28
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   28
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   28
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   28
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   27
               Left            =   8640
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   70
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   27
                  Left            =   -15
                  TabIndex        =   35
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   27
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇVì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   27
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   27
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   27
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   27
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   26
               Left            =   7380
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   26
                  Left            =   -15
                  TabIndex        =   34
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   26
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇUì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   26
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   26
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   26
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   26
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   25
               Left            =   6120
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   68
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   25
                  Left            =   -15
                  TabIndex        =   33
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   25
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇTì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   25
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   25
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   25
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   25
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   24
               Left            =   4860
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   67
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   24
                  Left            =   -15
                  TabIndex        =   32
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   24
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇSì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   24
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   24
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   24
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   24
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   23
               Left            =   3600
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   23
                  Left            =   -15
                  TabIndex        =   31
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   23
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇRì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   23
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   23
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   23
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   23
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   22
               Left            =   2340
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   22
                  Left            =   -15
                  TabIndex        =   30
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   22
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇQì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   22
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   22
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   22
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   22
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   21
               Left            =   1080
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   21
                  Left            =   -15
                  TabIndex        =   29
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   21
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇPì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   21
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   21
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   21
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   21
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   20
               Left            =   12420
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   20
                  Left            =   -15
                  TabIndex        =   28
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   20
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇOì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   20
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   20
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   20
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   20
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   19
               Left            =   11160
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   19
                  Left            =   -15
                  TabIndex        =   27
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   19
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇXì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   19
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   19
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   19
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   19
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   18
               Left            =   9900
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   18
                  Left            =   -15
                  TabIndex        =   26
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   18
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇWì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   18
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   18
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   18
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   18
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   17
               Left            =   8640
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   17
                  Left            =   -15
                  TabIndex        =   25
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   17
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇVì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   17
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   17
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   17
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   17
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   16
               Left            =   7380
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   16
                  Left            =   -15
                  TabIndex        =   24
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   16
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇU"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   16
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   16
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   16
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   16
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   15
               Left            =   6120
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   15
                  Left            =   -15
                  TabIndex        =   23
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   15
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇTì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   15
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   15
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   15
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   15
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   14
               Left            =   4860
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   14
                  Left            =   -15
                  TabIndex        =   22
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   14
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇSì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   14
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   14
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   14
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   14
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   13
               Left            =   3600
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   13
                  Left            =   -15
                  TabIndex        =   21
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   13
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇRì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   13
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   13
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   13
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   13
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   12
               Left            =   2340
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   12
                  Left            =   -15
                  TabIndex        =   20
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   12
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇQì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   12
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   12
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   12
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   12
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   11
               Left            =   1080
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   11
                  Left            =   -15
                  TabIndex        =   19
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   11
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇPì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   11
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   11
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   11
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   11
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   10
               Left            =   12420
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   10
                  Left            =   -15
                  TabIndex        =   18
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   10
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇOì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   10
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   10
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   10
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   10
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   9
               Left            =   11160
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   9
                  Left            =   -15
                  TabIndex        =   17
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   9
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇXì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   9
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   9
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   9
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   9
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   8
               Left            =   9900
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   8
                  Left            =   -15
                  TabIndex        =   16
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   8
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇWì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   8
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   8
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   8
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   8
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   7
               Left            =   8640
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   7
                  Left            =   -15
                  TabIndex        =   15
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   7
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇVì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   7
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   7
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   7
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   7
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   6
               Left            =   7380
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   6
                  Left            =   -15
                  TabIndex        =   14
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   6
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇUì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   6
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   6
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   6
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   6
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   5
               Left            =   6120
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   5
                  Left            =   -15
                  TabIndex        =   13
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   5
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇTì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   5
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   5
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   5
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   5
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   4
               Left            =   4860
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   4
                  Left            =   -15
                  TabIndex        =   12
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   4
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇSì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   4
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   4
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   4
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   4
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   3
               Left            =   3600
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   3
                  Left            =   -15
                  TabIndex        =   11
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   3
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇRì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   3
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   3
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   3
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   3
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   2
               Left            =   2340
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   2
                  Left            =   -15
                  TabIndex        =   10
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   2
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   2
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   2
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   2
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   2
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   12648384
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   1485
               Index           =   48
               Left            =   7785
               Top             =   4500
               Width           =   5850
               _ExtentX        =   10319
               _ExtentY        =   2619
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9.75
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
               cBoderWidth     =   2
               cAlingnment     =   4
               cBoderStyle     =   1
               Begin Cis3D_v60.CIS3D B1lb_LT 
                  Height          =   330
                  Left            =   1515
                  Top             =   240
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   582
                  BackColor       =   12648384
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -30
               End
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   1395
                  Index           =   49
                  Left            =   30
                  Top             =   60
                  Width           =   450
                  _ExtentX        =   794
                  _ExtentY        =   2461
                  BackColor       =   4210752
                  ForeColor       =   65535
                  Caption         =   "Ç‹ÇÈÇﬂèåè"
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
                  cBoderWidth     =   2
                  cFont3D         =   2
                  cPositionX      =   110
                  cPositionY      =   100
               End
               Begin Cis3D_v60.CIS3D CIS3D42 
                  Height          =   360
                  Left            =   615
                  Top             =   240
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   635
                  ForeColor       =   16711680
                  Caption         =   "ÿ∞ƒﬁ¿≤—"
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cAlingnment     =   1
                  cBoderStyle     =   2
               End
               Begin Cis3D_v60.CIS3D CIS3D43 
                  Height          =   390
                  Left            =   2865
                  Top             =   210
                  Width           =   2820
                  _ExtentX        =   4974
                  _ExtentY        =   688
                  ForeColor       =   16711680
                  Caption         =   "ójì˙éwíË"
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cAlingnment     =   1
                  cBoderStyle     =   2
                  Begin Cis3D_v60.CIS3D B1lb_Week 
                     Height          =   330
                     Index           =   1
                     Left            =   1035
                     Top             =   30
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   582
                     BackColor       =   12648384
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
                     cBoderColor1    =   8421504
                     cBoderColor2    =   -2147483643
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cAlingnment     =   7
                     cBoderStyle     =   1
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Week 
                     Height          =   330
                     Index           =   2
                     Left            =   1395
                     Top             =   30
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   582
                     BackColor       =   12648384
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
                     cBoderColor1    =   8421504
                     cBoderColor2    =   -2147483643
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cAlingnment     =   7
                     cBoderStyle     =   1
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Week 
                     Height          =   330
                     Index           =   3
                     Left            =   1755
                     Top             =   30
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   582
                     BackColor       =   12648384
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
                     cBoderColor1    =   8421504
                     cBoderColor2    =   -2147483643
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cAlingnment     =   7
                     cBoderStyle     =   1
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Week 
                     Height          =   330
                     Index           =   4
                     Left            =   2115
                     Top             =   30
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   582
                     BackColor       =   12648384
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
                     cBoderColor1    =   8421504
                     cBoderColor2    =   -2147483643
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cAlingnment     =   7
                     cBoderStyle     =   1
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Week 
                     Height          =   330
                     Index           =   5
                     Left            =   2475
                     Top             =   30
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   582
                     BackColor       =   12648384
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
                     cBoderColor1    =   8421504
                     cBoderColor2    =   -2147483643
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cAlingnment     =   7
                     cBoderStyle     =   1
                  End
               End
               Begin Cis3D_v60.CIS3D CIS3D44 
                  Height          =   390
                  Left            =   660
                  Top             =   930
                  Width           =   2295
                  _ExtentX        =   4048
                  _ExtentY        =   688
                  ForeColor       =   16711680
                  Caption         =   "î≠íç€Øƒ"
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cAlingnment     =   1
                  cBoderStyle     =   2
                  Begin Cis3D_v60.CIS3D B1lb_Lot 
                     Height          =   330
                     Left            =   855
                     Top             =   15
                     Width           =   1440
                     _ExtentX        =   2540
                     _ExtentY        =   582
                     BackColor       =   12648384
                     Caption         =   "999,999.000"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                     cPositionX      =   -20
                  End
               End
               Begin Cis3D_v60.CIS3D CIS3D45 
                  Height          =   375
                  Left            =   3375
                  Top             =   930
                  Width           =   2310
                  _ExtentX        =   4075
                  _ExtentY        =   661
                  ForeColor       =   16711680
                  Caption         =   "ì˙ït"
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cAlingnment     =   1
                  cBoderStyle     =   2
                  Begin Cis3D_v60.CIS3D B1lb_Day 
                     Height          =   330
                     Index           =   1
                     Left            =   510
                     Top             =   15
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   582
                     BackColor       =   12648384
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
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -20
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Day 
                     Height          =   330
                     Index           =   2
                     Left            =   870
                     Top             =   15
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   582
                     BackColor       =   12648384
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
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -20
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Day 
                     Height          =   330
                     Index           =   3
                     Left            =   1230
                     Top             =   15
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   582
                     BackColor       =   12648384
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
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -20
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Day 
                     Height          =   330
                     Index           =   4
                     Left            =   1590
                     Top             =   15
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   582
                     BackColor       =   12648384
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
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -20
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Day 
                     Height          =   330
                     Index           =   5
                     Left            =   1950
                     Top             =   15
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   582
                     BackColor       =   12648384
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
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -20
                  End
               End
               Begin VB.Shape Shape4 
                  Height          =   675
                  Left            =   510
                  Shape           =   4  'ä€Ç›ÇÃÇ†ÇÈí∑ï˚å`
                  Top             =   750
                  Width           =   2610
               End
               Begin VB.Shape Shape3 
                  Height          =   675
                  Left            =   3105
                  Shape           =   4  'ä€Ç›ÇÃÇ†ÇÈí∑ï˚å`
                  Top             =   750
                  Width           =   2715
               End
               Begin VB.Shape Shape2 
                  Height          =   675
                  Left            =   2595
                  Shape           =   4  'ä€Ç›ÇÃÇ†ÇÈí∑ï˚å`
                  Top             =   60
                  Width           =   3225
               End
               Begin VB.Shape Shape1 
                  Height          =   675
                  Left            =   510
                  Shape           =   4  'ä€Ç›ÇÃÇ†ÇÈí∑ï˚å`
                  Top             =   60
                  Width           =   2100
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   585
               Index           =   17
               Left            =   15
               Top             =   900
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   1032
               BackColor       =   16777152
               Caption         =   "ç°âÒ"
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
               cPositionX      =   30
               cPositionY      =   30
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   585
               Index           =   16
               Left            =   15
               Top             =   300
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   1032
               BackColor       =   12648384
               Caption         =   "ëOâÒ"
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
               cPositionX      =   30
               cPositionY      =   30
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Index           =   1
               Left            =   1080
               ScaleHeight     =   1470
               ScaleWidth      =   1245
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   1
                  Left            =   -15
                  TabIndex        =   9
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "#####0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  cILength        =   4
               End
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   300
                  Index           =   1
                  Left            =   -15
                  Top             =   -15
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  BackColor       =   8421631
                  ForeColor       =   16711680
                  Caption         =   "ÇPì˙"
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
                  cBoderWidth     =   2
                  cAlingnment     =   4
                  Begin Cis3D_v60.CIS3D B1lb_Kad 
                     Height          =   210
                     Index           =   1
                     Left            =   105
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr Çoñæí©"
                        Size            =   9.75
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
               Begin Cis3D_v60.CIS3D B1lb_NJSu 
                  Height          =   300
                  Index           =   1
                  Left            =   0
                  Top             =   1170
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LNJSu 
                  Height          =   300
                  Index           =   1
                  Left            =   0
                  Top             =   270
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16761024
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_LSu 
                  Height          =   300
                  Index           =   1
                  Left            =   0
                  Top             =   570
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   529
                  BackColor       =   16744703
                  Caption         =   "#####0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   315
               Index           =   0
               Left            =   -15
               Top             =   -15
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cBoderWidth     =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   285
               Index           =   1
               Left            =   330
               Top             =   300
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               ForeColor       =   16711680
               Caption         =   "ïKóvêî"
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
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   285
               Index           =   3
               Left            =   330
               Top             =   1200
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               ForeColor       =   16711680
               Caption         =   "ïKóvêî"
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
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   645
               Index           =   2
               Left            =   330
               Top             =   570
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   1138
               BackColor       =   16777215
               ForeColor       =   16711680
               Caption         =   "åvâÊêî"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9.75
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
               Height          =   585
               Index           =   19
               Left            =   15
               Top             =   2400
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   1032
               BackColor       =   16777152
               Caption         =   "ç°âÒ"
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
               cPositionX      =   30
               cPositionY      =   30
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   585
               Index           =   18
               Left            =   15
               Top             =   1800
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   1032
               BackColor       =   12648384
               Caption         =   "ëOâÒ"
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
               cPositionX      =   30
               cPositionY      =   30
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   315
               Index           =   4
               Left            =   -15
               Top             =   1500
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cBoderWidth     =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   285
               Index           =   5
               Left            =   330
               Top             =   1800
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               ForeColor       =   16711680
               Caption         =   "ïKóvêî"
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
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   285
               Index           =   7
               Left            =   330
               Top             =   2700
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               ForeColor       =   16711680
               Caption         =   "ïKóvêî"
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
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   645
               Index           =   6
               Left            =   330
               Top             =   2070
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   1138
               BackColor       =   16777215
               ForeColor       =   16711680
               Caption         =   "åvâÊêî"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9.75
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
               Height          =   585
               Index           =   21
               Left            =   15
               Top             =   3900
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   1032
               BackColor       =   16777152
               Caption         =   "ç°âÒ"
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
               cPositionX      =   30
               cPositionY      =   30
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   585
               Index           =   20
               Left            =   15
               Top             =   3300
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   1032
               BackColor       =   12648384
               Caption         =   "ëOâÒ"
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
               cPositionX      =   30
               cPositionY      =   30
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   315
               Index           =   8
               Left            =   -15
               Top             =   3000
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cBoderWidth     =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   285
               Index           =   9
               Left            =   330
               Top             =   3300
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               ForeColor       =   16711680
               Caption         =   "ïKóvêî"
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
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   285
               Index           =   11
               Left            =   330
               Top             =   4200
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               ForeColor       =   16711680
               Caption         =   "ïKóvêî"
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
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   645
               Index           =   10
               Left            =   330
               Top             =   3570
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   1138
               BackColor       =   16777215
               ForeColor       =   16711680
               Caption         =   "åvâÊêî"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9.75
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
               Height          =   585
               Index           =   23
               Left            =   15
               Top             =   5415
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   1032
               BackColor       =   16777152
               Caption         =   "ç°âÒ"
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
               cPositionX      =   30
               cPositionY      =   30
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   600
               Index           =   22
               Left            =   0
               Top             =   4800
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   1058
               BackColor       =   12648384
               Caption         =   "ëOâÒ"
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
               cPositionX      =   30
               cPositionY      =   30
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   315
               Index           =   12
               Left            =   -15
               Top             =   4500
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cBoderWidth     =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   285
               Index           =   13
               Left            =   330
               Top             =   4800
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               ForeColor       =   16711680
               Caption         =   "ïKóvêî"
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
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   285
               Index           =   15
               Left            =   330
               Top             =   5715
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               ForeColor       =   16711680
               Caption         =   "ïKóvêî"
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
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   645
               Index           =   14
               Left            =   330
               Top             =   5070
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   1138
               BackColor       =   16777215
               ForeColor       =   16711680
               Caption         =   "åvâÊêî"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9.75
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
            Begin Cis3D_v60.CIS3D B1lb_NYM 
               Height          =   315
               Index           =   2
               Left            =   5940
               Top             =   4500
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   556
               BackColor       =   16777088
               ForeColor       =   16711680
               Caption         =   "2004îN99åé"
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
               cBoderWidth     =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D B1lb_LGNJSu 
               Height          =   285
               Index           =   2
               Left            =   5940
               Top             =   4800
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   503
               BackColor       =   12648384
               Caption         =   "##,###,##0.000"
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
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
            Begin Cis3D_v60.CIS3D CIS3D47 
               Height          =   630
               Left            =   5940
               Top             =   5370
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   1111
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cBoderStyle     =   1
               cPositionY      =   30
               Begin Cis3D_v60.CIS3D B1lb_GNJSu 
                  Height          =   300
                  Index           =   2
                  Left            =   0
                  Top             =   315
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "##,###,##0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_GSu 
                  Height          =   300
                  Index           =   2
                  Left            =   0
                  Top             =   0
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   529
                  BackColor       =   12640511
                  ForeColor       =   0
                  Caption         =   "##,###,##0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D32 
               Height          =   630
               Left            =   4140
               Top             =   5370
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   1111
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cBoderStyle     =   1
               cPositionY      =   30
               Begin Cis3D_v60.CIS3D B1lb_GNJSu 
                  Height          =   300
                  Index           =   1
                  Left            =   0
                  Top             =   315
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "##,###,##0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_GSu 
                  Height          =   300
                  Index           =   1
                  Left            =   0
                  Top             =   0
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   529
                  BackColor       =   12640511
                  ForeColor       =   0
                  Caption         =   "##,###,##0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin Cis3D_v60.CIS3D B1lb_LGNJSu 
               Height          =   285
               Index           =   1
               Left            =   4140
               Top             =   4800
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   503
               BackColor       =   12648384
               Caption         =   "##,###,##0.000"
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
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
            Begin Cis3D_v60.CIS3D B1lb_NYM 
               Height          =   315
               Index           =   1
               Left            =   4140
               Top             =   4500
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   556
               BackColor       =   16777088
               ForeColor       =   16711680
               Caption         =   "2004îN99åé"
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
               cBoderWidth     =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D B1lb_LGSu 
               Height          =   285
               Index           =   1
               Left            =   4140
               Top             =   5085
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   503
               BackColor       =   12648384
               Caption         =   "##,###,##0.000"
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
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
            Begin Cis3D_v60.CIS3D B1lb_LGSu 
               Height          =   285
               Index           =   2
               Left            =   5940
               Top             =   5085
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   503
               BackColor       =   12648384
               Caption         =   "##,###,##0.000"
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
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
            Begin Cis3D_v60.CIS3D CIS3D8 
               Height          =   630
               Left            =   2340
               Top             =   5370
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   1111
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cBoderStyle     =   1
               cPositionY      =   30
               Begin Cis3D_v60.CIS3D B1lb_GNJSu 
                  Height          =   300
                  Index           =   0
                  Left            =   0
                  Top             =   315
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   529
                  BackColor       =   16777215
                  ForeColor       =   0
                  Caption         =   "##,###,##0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_GSu 
                  Height          =   300
                  Index           =   0
                  Left            =   0
                  Top             =   0
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   529
                  BackColor       =   12640511
                  ForeColor       =   0
                  Caption         =   "##,###,##0.000"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin Cis3D_v60.CIS3D B1lb_LGNJSu 
               Height          =   285
               Index           =   0
               Left            =   2340
               Top             =   4800
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   503
               BackColor       =   12648384
               Caption         =   "##,###,##0.000"
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
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
            Begin Cis3D_v60.CIS3D B1lb_NYM 
               Height          =   315
               Index           =   0
               Left            =   2340
               Top             =   4500
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   556
               BackColor       =   16777088
               ForeColor       =   16711680
               Caption         =   "2004îN99åé"
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
               cBoderWidth     =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D B1lb_LGSu 
               Height          =   285
               Index           =   0
               Left            =   2340
               Top             =   5085
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   503
               BackColor       =   12648384
               Caption         =   "##,###,##0.000"
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
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
         End
      End
   End
End
Attribute VB_Name = "CNJZ0030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :   çﬁóøî≠íçì‡é¶ì¸óÕ
'**       ÉtÉHÅ[ÉÄID    :   CNJZ0030
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :   2004/03/22  By CIS
'**       ïœ  çX  ì˙    :   2004/04/09  ëOâÒÉeÅ[ÉuÉãçÏê¨(ïœçXï™Å®ëSÇƒ)
'**       ïœ  çX  ì˙    :   2005/04/05  ëŒè€îNåéÇÃ¡™Ø∏í«â¡Å®ì‡é¶îNåéÇ…ëŒÇµÇƒÅAìñåéÅEóÇåéÅEóÇÅXåéï™ÇÃÇRÉñåéÇÃÇ›ÇnÇjÇ∆Ç∑ÇÈÅB
'**       ïœ  çX  ì˙    :   2010/02/23  ëŒè€îNåééZèoÇÃïœçXÅ®çﬁóøî≠íçä«óù_ì‡é¶ëŒè€ÉfÅ[É^(CISKANRIÅF'Option'_11)Ç™ìñåé(0)ÇÃèÍçá
'**                                     CisFun.Date_AddÇâÒîÇ∑ÇÈÅBì‡é¶îNåéÇÇªÇÃÇ‹Ç‹ëŒè€îNåéÇ∆Ç∑ÇÈÅB
'**
'   200312,200312,000001111123,0.00,0.00,0.00,0001A,01
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' Ã´∞—±∏√®Ãﬁ Ã◊∏ﬁ
    Dim ProcHB                  As String * 2   ' ì¸óÕàÊêÿë÷
    Dim SyoriKB                 As String       ' èàóùãÊï™(êVãK•èCê≥•çÌèú ìô)
    Dim SyoriNM                 As String       ' èàóùñºèÃ(êVãK•èCê≥•çÌèú ìô)
    Dim YMAdd                   As Byte         ' ì‡é¶ëŒè€îNåé(â¡éZíl)
    Dim hd_ZaiKnrNo             As Long         ' çﬁóøä«óùî‘çÜ  ( DBInputÇ≈ÉZÉbÉg )
    Dim hd_SunpKB               As Byte         ' ê°ñ@ãÊï™      ( DBInputÇ≈ÉZÉbÉg )

    Dim mOndayColor             As Long         ' â“ì≠ì˙
    Dim mOffdayColor            As Long         ' ãxì˙
    Dim mNInitColor             As Long         ' ì‡é¶
    Dim mNHenColor              As Long         ' ì‡é¶(ïœçX)
    Dim mHHenColor              As Long         ' çﬁóøî≠íçì‡é¶(ïœçX)

    Dim wShift                  As Integer      ' PageUp/PageDownéû ( ëŒè€îNåé )
    Dim sv_Choku                As String       ' íºëóêÊ(ëﬁîóp)
    Dim sv_Zaist                As String       ' çﬁéø(ëﬁî)
    Dim sv_Itatu                As Currency     ' î¬å˙(ëﬁî)
    Dim sv_Width                As Currency     ' ïùÅ@(ëﬁî)
    Dim sv_Long                 As Currency     ' í∑Ç≥(ëﬁî)
    Dim sv_NJYM                 As String       ' ì‡é¶îNåé
    
    Dim wInt                    As Integer

'( ÉvÉçÉOÉâÉÄòAågóp )
    Dim cmd_KB                  As Byte
    Dim cmd_NJYM                As String
    Dim cmd_YM                  As String
    Dim cmd_Zaist               As String
    Dim cmd_Itatu               As Currency
    Dim cmd_Width               As Currency
    Dim cmd_Long                As Currency
    Dim cmd_Torcd               As String
    Dim cmd_Ukeir               As String
    Dim wSplit()                As String
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
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    NaiHatu_Back(0).Visible = False
    NaiHatu_Back(1).Visible = False
    NaiHatu_Back(2).Visible = False
    B1lb_HCSu.Visible = False
    B1lb_SKKbn.Move NaiHatu_Back(1).Left, 390
    B1lb_NaHaKB.Move NaiHatu_Back(2).Left, 390
    
    Timer1.Interval = 1000   ' éûçèï\é¶ä‘äu
    Timer1.Enabled = True    ' éûçèï\é¶äJén
    Timer2.Interval = 500    ' èàóùãÊï™ì_ñ≈ä‘äu
    Timer2.Enabled = False   ' èàóùãÊï™ì_ñ≈ã÷é~
    
'   #-------------------#
'   #  çÄ ñ⁄ èâ ä˙ âª    #
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
    ProcHB = "H1"
    Call DispChange(ProcHB)
    
'< ÉpÉâÉÅÅ[É^É`ÉFÉbÉN >
    If RTrim$(Command) <> "" Then
'       ( ì¡íËÉfÅ[É^ÇÃÇ›ÇÃÉÅÉìÉe )
        cmd_KB = 1
        wSplit = Split(Command, ",")

        If UBound(wSplit) <> 7 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ÉpÉâÉÅÅ[É^êîÇ™ê≥ÇµÇ≠Ç†ÇËÇ‹ÇπÇÒ            "
                .MB_Button = Error
                .MBOX
            End With
            Call PB_END_Click
            Exit Sub
        End If
    
        cmd_NJYM = RTrim$(wSplit(0))
        cmd_YM = RTrim$(wSplit(1))
        cmd_Zaist = RTrim$(wSplit(2))
        cmd_Itatu = CisFun.Val2(wSplit(3))
        cmd_Width = CisFun.Val2(wSplit(4))
        cmd_Long = CisFun.Val2(wSplit(5))
        cmd_Torcd = RTrim$(wSplit(6))
        cmd_Ukeir = RTrim$(wSplit(7))

'       ( èàóùãÊï™èâä˙ï\é¶ )
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
        
        PB_END.BtnCaption = "ñﬂÇÈ"
    Else
        cmd_KB = 0

'       ( èàóùãÊï™èâä˙ï\é¶ )
        SyoriKB = "ADD"
        Call SyoriName(SyoriKB)
    End If

'   #-------------------#
'   # âÊñ íÜâõà íuï\é¶   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    ÅsÇ`ÇÉÇîÇâÇñÇÅÇîÇÖ ÉCÉxÉìÉgÅt                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    FormAct = Not FormAct

' ì‡é¶ëŒè€îNåé
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 11
    Call IRNRead("", 1)
    YMAdd = CisFun.Val2(IRN.oyian)
    
    sv_Zaist = "": sv_Itatu = 0: sv_Width = 0:  sv_Long = 0

'< î≠íçì‡é¶îNåé >
    Call GetNJProc_YM(3)
    If RTrim$(gn_HNYM) = "" Then gn_HNYM = Format(Now(), "YYYYMM")
    iH1_NJYM = gn_HNYM
    Call iH1_NJYM_LostFocus

'( Back Color )
    mOffdayColor = BLB_DD(1).BackColor
    mOndayColor = BLB_DD(2).BackColor
    BLB_DD(1).BackColor = mOndayColor
    
    mNHenColor = B1lb_LNJSu(1).BackColor
    mNInitColor = B1lb_LNJSu(2).BackColor
    B1lb_LNJSu(1).BackColor = mNInitColor
    
    mHHenColor = B1lb_LSu(1).BackColor
    B1lb_LSu(1).BackColor = mNInitColor

'( ójì˙ÉZÉbÉg )
    Call SetWeek(iH1_YM, "", "")

    If cmd_KB = 1 Then
'       ( ì¡íËÉfÅ[É^ÇÃÇ›ÇÃÉÅÉìÉe )
        iH1_NJYM = cmd_NJYM
        iH1_YM = cmd_YM
        iH1_Zaist = cmd_Zaist
        iH1_Itatu = Format(cmd_Itatu, "###0.00;###0.00;#")
        iH1_Width = Format(cmd_Width, "###0.00;###0.00;#")
        iH1_Long = Format(cmd_Long, "###0.00;###0.00;#")
        iH1_Torcd = cmd_Torcd
        iH1_Ukeir = cmd_Ukeir
    
        Call ReturnPress
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
           Case vbKeyF1:        If PB_ADD.Visible Then Call PB_ADD_Click    'ÅyêVãKÅz
           Case vbKeyF2:        If PB_MNT.Visible Then Call PB_MNT_Click    'ÅyèCê≥Åz
           Case vbKeyF3:        If PB_DEL.Visible Then Call PB_DEL_Click    'ÅyçÌèúÅz
           Case vbKeyF5:        If PB_REP.Visible Then Call PB_REP_Click    'ÅyïœçXëOÅz
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  'ÅyåüçıÅz
           Case vbKeyPageUp:    If PB_BACK.Visible Then wShift = Shift: Call PB_Back_Click  'Åyå„ëﬁÅz
           Case vbKeyPageDown:  If PB_NEXT.Visible Then wShift = Shift: Call PB_Next_Click  'ÅyëOêiÅz
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click    'ÅyéÊè¡Åz
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    'ÅyèIóπÅz
           Case vbKeyReturn
                                If PB_ENT.Visible Then                      'Åyì¸óÕÅz
                                    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                    Call PB_ENT_Click
                                    Dummy.Enabled = False
                                End If
           Case Else:           Exit Function
    End Select
    
    Key_Acc = True
End Function
'+---------------------+
'+    èI  óπ  ÉL Å[     +
'+---------------------+
Private Sub PB_END_Click()
    Call CisDB.DBDISConnect                         ' √ﬁ∞¿Õﬁ∞Ωêÿíf
    Call CisFrm.UnLoadDisp                          ' Ã´∞—µÃﬁºﬁ™∏ƒè¡ãé
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
        If cmd_KB = 1 Then
'           ( ì¡íËÉfÅ[É^ÇÃÇ›ÇÃÉÅÉìÉe )
            With CisFun
                .MB_Lines = 4
                .MB_MSG(1) = "    èàóùÇèIóπÇµÇ‹Ç∑            "
                .MB_MSG(3) = "    ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH            "
                .MB_Button = Yes_No
                If .MBOX Then Call PB_END_Click
            End With
        Else
            ProcHB = "H1"
            Call DispChange(ProcHB)
            Call HeadBodyClear("B")
            Call SyoriName(SyoriKB)
            iH1_YM.SetFocus
        End If
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
    
    If wShift = 0 Then
        gSL_Select = "Select Distinct ì‡é¶îNåé, åvâÊîNåé, çﬁéø, î¬å˙, ïù, í∑Ç≥, çﬁóøä«óùî‘çÜ, ê°ñ@ãÊï™,"
        gSL_Select = gSL_Select & " éËîzêÊ, éÛì¸" & vbCrLf
        gSL_Select = gSL_Select & " From çﬁóøî≠íçì‡é¶ÉeÅ[ÉuÉã" & vbCrLf
        gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   SubString( çﬁéø   + Space(25), 1, 25 )" & vbCrLf
        gSL_Select = gSL_Select & "     + Space( 7 - Len( î¬å˙ ) ) + Cast( î¬å˙ As VarChar(7) )" & vbCrLf
        gSL_Select = gSL_Select & "     + Space( 7 - Len( ïù )   ) + Cast( ïù   As VarChar(7) )" & vbCrLf
        gSL_Select = gSL_Select & "     + Space( 7 - Len( í∑Ç≥ ) ) + Cast( í∑Ç≥ As VarChar(7) )" & vbCrLf
        gSL_Select = gSL_Select & "     + SubString( éËîzêÊ + Space(7),  1,  7 )" & vbCrLf
        gSL_Select = gSL_Select & "     + SubString( éÛì¸   + Space(2),  1,  2 )" & vbCrLf
        gSL_Select = gSL_Select & "       < '"
        gSL_Select = gSL_Select & RTrim$(iH1_Zaist) & Space(25 - Len(RTrim$(iH1_Zaist)))
        gSL_Select = gSL_Select & CisFun.RSetFld(CisFun.Val2(iH1_Itatu), 7, "###0.00")
        gSL_Select = gSL_Select & CisFun.RSetFld(CisFun.Val2(iH1_Width), 7, "###0.00")
        gSL_Select = gSL_Select & CisFun.RSetFld(CisFun.Val2(iH1_Long), 7, "###0.00")
        gSL_Select = gSL_Select & RTrim$(iH1_Torcd) & Space(7 - Len(RTrim$(iH1_Torcd)))
        gSL_Select = gSL_Select & RTrim$(iH1_Ukeir) & Space(2 - Len(RTrim$(iH1_Ukeir))) & "'" & vbCrLf
        gSL_Select = gSL_Select & " ORDER BY çﬁéø Desc, î¬å˙ Desc, ïù Desc, í∑Ç≥ Desc, éËîzêÊ Desc, éÛì¸ Desc" & vbCrLf
    Else
        gSL_Select = "Select Distinct ì‡é¶îNåé, åvâÊîNåé, çﬁéø, î¬å˙, ïù, í∑Ç≥, çﬁóøä«óùî‘çÜ, ê°ñ@ãÊï™,"
        gSL_Select = gSL_Select & " éËîzêÊ, éÛì¸" & vbCrLf
        gSL_Select = gSL_Select & " From çﬁóøî≠íçì‡é¶ÉeÅ[ÉuÉã" & vbCrLf
        gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   çﬁéø = '" & RTrim$(iH1_Zaist) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   î¬å˙ = '" & CisFun.Val2(iH1_Itatu) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   ïù = '" & CisFun.Val2(iH1_Width) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   í∑Ç≥ = '" & CisFun.Val2(iH1_Long) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim$(iH1_Torcd) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim$(iH1_Ukeir) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   åvâÊîNåé < '" & RTrim$(iH1_YM) & "'" & vbCrLf
        gSL_Select = gSL_Select & " ORDER BY åvâÊîNåé Desc" & vbCrLf
    End If

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
    
    If wShift = 0 Then
        gSL_Select = "Select Distinct ì‡é¶îNåé, åvâÊîNåé, çﬁéø, î¬å˙, ïù, í∑Ç≥, çﬁóøä«óùî‘çÜ, ê°ñ@ãÊï™,"
        gSL_Select = gSL_Select & " éËîzêÊ, éÛì¸" & vbCrLf
        gSL_Select = gSL_Select & " From çﬁóøî≠íçì‡é¶ÉeÅ[ÉuÉã" & vbCrLf
        gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   SubString( çﬁéø   + Space(25), 1, 25 )" & vbCrLf
        gSL_Select = gSL_Select & "     + Space( 7 - Len( î¬å˙ ) ) + Cast( î¬å˙ As VarChar(7) )" & vbCrLf
        gSL_Select = gSL_Select & "     + Space( 7 - Len( ïù )   ) + Cast( ïù   As VarChar(7) )" & vbCrLf
        gSL_Select = gSL_Select & "     + Space( 7 - Len( í∑Ç≥ ) ) + Cast( í∑Ç≥ As VarChar(7) )" & vbCrLf
        gSL_Select = gSL_Select & "     + SubString( éËîzêÊ + Space(7),  1,  7 )" & vbCrLf
        gSL_Select = gSL_Select & "     + SubString( éÛì¸   + Space(2),  1,  2 )" & vbCrLf
        gSL_Select = gSL_Select & "       > '"
        gSL_Select = gSL_Select & RTrim$(iH1_Zaist) & Space(25 - Len(RTrim$(iH1_Zaist)))
        gSL_Select = gSL_Select & CisFun.RSetFld(CisFun.Val2(iH1_Itatu), 7, "###0.00")
        gSL_Select = gSL_Select & CisFun.RSetFld(CisFun.Val2(iH1_Width), 7, "###0.00")
        gSL_Select = gSL_Select & CisFun.RSetFld(CisFun.Val2(iH1_Long), 7, "###0.00")
        gSL_Select = gSL_Select & RTrim$(iH1_Torcd) & Space(7 - Len(RTrim$(iH1_Torcd)))
        gSL_Select = gSL_Select & RTrim$(iH1_Ukeir) & Space(2 - Len(RTrim$(iH1_Ukeir))) & "'" & vbCrLf
        gSL_Select = gSL_Select & " ORDER BY çﬁéø, î¬å˙, ïù, í∑Ç≥, éËîzêÊ, éÛì¸" & vbCrLf
    Else
        gSL_Select = "Select Distinct ì‡é¶îNåé, åvâÊîNåé, çﬁéø, î¬å˙, ïù, í∑Ç≥, çﬁóøä«óùî‘çÜ, ê°ñ@ãÊï™,"
        gSL_Select = gSL_Select & " éËîzêÊ, éÛì¸" & vbCrLf
        gSL_Select = gSL_Select & " From çﬁóøî≠íçì‡é¶ÉeÅ[ÉuÉã" & vbCrLf
        gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   çﬁéø = '" & RTrim$(iH1_Zaist) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   î¬å˙ = '" & CisFun.Val2(iH1_Itatu) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   ïù = '" & CisFun.Val2(iH1_Width) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   í∑Ç≥ = '" & CisFun.Val2(iH1_Long) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim$(iH1_Torcd) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim$(iH1_Ukeir) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   åvâÊîNåé > '" & RTrim$(iH1_YM) & "'" & vbCrLf
        gSL_Select = gSL_Select & " ORDER BY åvâÊîNåé" & vbCrLf
    End If

    Call DBInput("Next")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_YM.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    ïœçXëO (F5)
'+----------------------------+
Private Sub PB_REP_Click()
    wInt = 0
    On Error Resume Next: wInt = ActiveControl.Index: On Error GoTo 0
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
        
    For gInt = 1 To 31
        iB1_Su(gInt) = CisFun.RSetFld(iB1_Su(gInt).Tag, 10, "##0.000;-##0.000;#")
    Next gInt
    B1lb_GSu(0) = CisFun.RSetFld(B1lb_GSu(0).Tag, 11, "#,##0.000;-#,##0.000;#")
    
    If wInt = 0 Then wInt = 1
    iB1_Su(wInt).SetFocus
    
    Dummy.Enabled = False
End Sub
'+----------------------------+
'+    åü  çı  ÉL Å[(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Zaist" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Itatu" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Width" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Long" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Tehacd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_Tehacd: Exit Sub
End Sub
'+-----------------------------------+
'+       çﬁóøåüçı                     +
'+-----------------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RV_Zaist = RTrim$(iH1_Zaist)
    
    CKK0070.Show vbModal
    Unload CKK0070
    Set CKK0070 = Nothing

    If RV_Rtn Then
        iH1_Zaist = RTrim$(RV_Zaist)
        iH1_Itatu = Format(RV_Itatu, "###0.00;###0.00;#")
        iH1_Width = Format(RV_Width, "###0.00;###0.00;#")
        iH1_Long = Format(RV_Long, "###0.00;###0.00;#")
    End If
End Sub
'+-----------------------------------+
'+       éËîzêÊåüçı
'+-----------------------------------+
Private Sub Look_Tehacd()
    RV_Left = 0
    RV_Top = 0
    RV_NJYM = RTrim(iH1_NJYM)
    RV_YM = RTrim(iH1_YM)
    RV_Zaist = RTrim$(iH1_Zaist)
    RV_Itatu = CisFun.Val2(iH1_Itatu)
    RV_Width = CisFun.Val2(iH1_Width)
    RV_Long = CisFun.Val2(iH1_Long)
    If SyoriKB = "ADD" Then RV_Kubun = "MST" Else RV_Kubun = "ZNJ"
    
    CNJZ0025.Show vbModal
    Unload CNJZ0025
    Set CNJZ0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_Torcd
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RV_Ukeir
    End If
End Sub
'+-----------------------------+
'+  iH1_NJYM    ì‡é¶îNåé
'+-----------------------------+
Private Sub iH1_NJYM_GotFocus()
    sv_NJYM = RTrim$(iH1_NJYM)
End Sub
Private Sub iH1_NJYM_LostFocus()
    If RTrim$(iH1_NJYM) <> sv_NJYM Then
       '## CHG-S 2010/02/23 ##
'        With CisFun
'            .Date_RsFormat = fYM
'            .Date_Add iH1_NJYM, M, CInt(YMAdd)
'
'            iH1_YM = .Date_Result
'        End With
       
        If YMAdd = 0 Then
            iH1_YM = RTrim$(iH1_NJYM)
        Else
            With CisFun
                .Date_RsFormat = fYM
                .Date_Add iH1_NJYM, M, CInt(YMAdd)
        
                iH1_YM = .Date_Result
            End With
        End If
       '## CHG-E 2010/02/23 ##
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
'+  iH1_Itatu
'+-----------------------------+
Private Sub iH1_Itatu_GotFocus()
    PB_Look.Tag = ActiveControl.Name:   PB_Look.Visible = True
End Sub
Private Sub iH1_Itatu_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_Width
'+-----------------------------+
Private Sub iH1_Width_GotFocus()
    PB_Look.Tag = ActiveControl.Name:   PB_Look.Visible = True
End Sub
Private Sub iH1_Width_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_Long
'+-----------------------------+
Private Sub iH1_Long_GotFocus()
    PB_Look.Tag = ActiveControl.Name:   PB_Look.Visible = True
End Sub
Private Sub iH1_Long_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    If (RTrim$(iH1_Zaist) <> sv_Zaist Or _
       CisFun.Val2(iH1_Itatu) <> sv_Itatu Or _
       CisFun.Val2(iH1_Width) <> sv_Width Or _
       CisFun.Val2(iH1_Long) <> sv_Long) And _
       ActiveControl.Name <> "Dummy" And _
       ActiveControl.Name <> "iH1_YM" Then
        Call SetTorcd
        sv_Zaist = RTrim$(iH1_Zaist)
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
        H1lb_Tornm = RTrim(TRM_RName)
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
'+  iB1_Su
'+-----------------------------+
Private Sub iB1_Su_LostFocus(Index As Integer)
    gLong = 0
    For gInt = 1 To 31
        gLong = gLong + CisFun.Val2(iB1_Su(gInt))
    Next gInt
    B1lb_GSu(0) = Format(gLong, "#,##0.000;-#,##0.000;#")
End Sub
'**************************************
'*                                    *
'*      èàóùãÊï™Ç…ÇÊÇËãÊï™ñºï\é¶        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
    Static w_Color      As Long
    w_Color = &H80000008 'èâä˙êFæØƒ(çï)
    
    Select Case SyoriKB
        Case "ADD"
                            SyoriNM = "í«â¡"
                            K_Sykbnm.ForeColor = &HFFFF&            '(â©êF)
        Case "MNT"
                            SyoriNM = "èCê≥"
                            K_Sykbnm.ForeColor = &HFF0000           '(ê¬êF)
        Case "DEL"
                            SyoriNM = "çÌèú"
                            K_Sykbnm.ForeColor = &HFF&              '(ê‘êF)
                            w_Color = &HFF&
        Case Else:          SyoriNM = ""
    End Select
    
    For Each gObj In Me.Controls
        If gObj.Name Like "i*" Then gObj.ForeColor = w_Color
        If gObj.Name Like "*Comb*" Then gObj.ForeColor = w_Color
    
'       * PF-Key ï\é¶ê›íË *
        If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
           gObj.Name = "PB_DEL" Then
                gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
        End If
    Next gObj
    
    Timer2.Enabled = False
    K_Sykbnm = SyoriNM
    K_Sykbnm.Visible = Not (SyoriNM = "")
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
    If RTrim$(K_Sykbnm) <> "" Then K_Sykbnm = "" Else K_Sykbnm = SyoriNM
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
    
    H1Mode = False:     H1Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    
'   ¿ﬁ–∞çÄñ⁄(âÊñ ÇÃÉ`ÉâÉcÉLñhé~)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "B1":      B1Mode = True:  B1Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    
    PB_Look.Visible = H1Mode
    PB_ADD.Visible = H1Mode
    PB_MNT.Visible = H1Mode
    PB_DEL.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
' *CHG* 2004/04/07
'    PB_CAN.Visible = Not H1Mode
'    PB_END.Visible = H1Mode
    If cmd_KB = 0 Then
        PB_CAN.Visible = Not H1Mode
    End If
    
    PB_Look.Visible = False
    PB_REP.Visible = B1Mode
    
'   ¿ﬁ–∞çÄñ⁄(ì¸óÕã÷é~)
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
        For gLong = 1 To 31
            B1lb_LSu(gLong).BackColor = mNInitColor
            B1lb_LNJSu(gLong).BackColor = mNInitColor
            B1lb_LSu(gLong).Tag = ""
            iB1_Su(gLong).Tag = ""
        Next gLong
    
        NaiHatu_Back(1) = ""
        NaiHatu_Back(2) = ""
    End If
End Sub
'****************************
'*      ì¸óÕèàóù(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "H1" Then
        If cmd_KB = 1 Then
            If Not Head1Chk_Cmd Then
                Call PB_END_Click
                GoTo ReturnPress_Ed
            End If
        Else
            If Not Head1Chk Then GoTo ReturnPress_Ed
        End If

'       << çÌèúèàóù >>
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
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
            
        If cmd_KB = 1 Then
'           ( ì¡íËÉfÅ[É^ÇÃÇ›ÇÃÉÅÉìÉe )
            Call PB_END_Click
        Else
            ProcHB = "H1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
        End If
        GoTo ReturnPress_Ed
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'*----------------------------------------------------------------*
'*
'*                  HEAD1¡™Ø∏èàóù   ( í èÌì¸óÕéû )
'*
'*----------------------------------------------------------------*
Private Function Head1Chk() As Boolean
    Dim wYm     As String                       '2005.04.05 Add
    Head1Chk = False
    
' * ì‡é¶îNåé
    If Trim(iH1_NJYM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    ëŒè€îNåéÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_NJYM.SetFocus
       Exit Function
    End If
' * ëŒè€îNåé
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
'Å§Å§---------------------------------------------<< 2005/04/05 >> Add Start
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add iH1_NJYM, M, 2
        wYm = .Date_Result
    End With
    If iH1_YM < iH1_NJYM Or iH1_YM > wYm Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    ëŒè€îNåéÇÕÅAÅuì‡é¶îNåéÅ`ì‡é¶îNåéÇÃÇQÉïåéå„ÅvÇÃîÕàÕÇ≈ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢    "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_YM.SetFocus
       Exit Function
    End If
        
'Å¢Å¢---------------------------------------------<< 2005/04/05 >> Add End
' * ïiî‘
    If Trim(iH1_Zaist) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    çﬁéøÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Zaist.SetFocus
       Exit Function
    End If

' * édì¸êÊ é©ìÆï\é¶
    If RTrim$(iH1_Torcd) = "" Then Call SetTorcd

' * édì¸êÊ
    If Trim(iH1_Torcd) <> "" Then
        gSL_Select = " SELECT * "
        gSL_Select = gSL_Select & " FROM éÊà¯êÊÉ}ÉXÉ^ "
        gSL_Select = gSL_Select & " WHERE éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   And éÊà¯êÊãÊï™ = 1"
        gSL_Select = gSL_Select & "   And é–ì‡ãÊï™ = 0"
        If Not TRMRead(gSL_Select, 1) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    édì¸êÊ(éÊà¯êÊÉ}ÉXÉ^ñ¢ìoò^)            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_Torcd.SetFocus
           Exit Function
        End If
    End If
' * éÛì¸
    If Trim(iH1_Ukeir) <> "" Then
        gSL_Select = " SELECT * "
        gSL_Select = gSL_Select & " FROM éÊà¯êÊéÛì¸É}ÉXÉ^ "
        gSL_Select = gSL_Select & " Where éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & " And   éÊà¯êÊãÊï™ = 1"
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
    End If
    
' * çﬁóøéÊà¯êÊÉ}ÉXÉ^
    gSL_Select = "Select çﬁéø From çﬁóøéÊà¯êÊÉ}ÉXÉ^"
    gSL_Select = gSL_Select & " Where çﬁéø = '" & RTrim$(iH1_Zaist) & "'"
    gSL_Select = gSL_Select & " And   î¬å˙ =  " & CisFun.Val2(iH1_Itatu)
    gSL_Select = gSL_Select & " And    ïù  =  " & CisFun.Val2(iH1_Width)
    gSL_Select = gSL_Select & " And   í∑Ç≥ =  " & CisFun.Val2(iH1_Long)
    gSL_Select = gSL_Select & " And   éÊà¯êÊ = '" & RTrim$(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim$(iH1_Ukeir) & "'"
    If Not ZTMRead(gSL_Select, 1) Then
         With CisFun
             .MB_Lines = 4
             .MB_MSG(2) = "    çﬁóøéÊà¯êÊÉ}ÉXÉ^ñ¢ìoò^            "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
         End With
        iH1_Zaist.SetFocus
        Exit Function
    End If

    gSL_Select = "Select * From çﬁóøî≠íçì‡é¶ÉeÅ[ÉuÉã"
    gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & " And   çﬁéø = '" & RTrim(iH1_Zaist) & "'"
    gSL_Select = gSL_Select & " And   î¬å˙ =  " & CisFun.Val2(iH1_Itatu)
    gSL_Select = gSL_Select & " And    ïù  =  " & CisFun.Val2(iH1_Width)
    gSL_Select = gSL_Select & " And   í∑Ç≥ =  " & CisFun.Val2(iH1_Long)
    gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    If DBInput("INV") Then
       If SyoriKB = "ADD" Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    çﬁóøî≠íçì‡é¶ÉfÅ[É^ ìoò^çœ            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          iH1_YM.SetFocus
          Exit Function
       End If
    Else
       If SyoriKB <> "ADD" Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    çﬁóøî≠íçì‡é¶ÉfÅ[É^ ñ¢ìoò^            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          iH1_YM.SetFocus
          Exit Function
       End If
    End If
    
    Head1Chk = True
End Function
'*----------------------------------------------------------------*
'*
'*      HEAD1¡™Ø∏èàóù   ( ÉvÉçÉOÉâÉÄòAågéûÇÃÉwÉbÉhÉ`ÉFÉbÉN )
'*
'*----------------------------------------------------------------*
Private Function Head1Chk_Cmd() As Boolean
    Head1Chk_Cmd = False
    
' * ì‡é¶îNåé
    If Trim(iH1_NJYM) = "" Then GoTo Head1Chk_Cmd_Err

' * ëŒè€îNåé
    If Trim(iH1_YM) = "" Then GoTo Head1Chk_Cmd_Err

' * çﬁéø
    If Trim(iH1_Zaist) = "" Then GoTo Head1Chk_Cmd_Err

' * éËîzêÊ
    If Trim(iH1_Torcd) <> "" Then
        gSL_Select = " SELECT * "
        gSL_Select = gSL_Select & " FROM éÊà¯êÊÉ}ÉXÉ^ "
        gSL_Select = gSL_Select & " WHERE éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   And éÊà¯êÊãÊï™ = 1"
        gSL_Select = gSL_Select & "   And é–ì‡ãÊï™ = 0"
        If Not TRMRead(gSL_Select, 1) Then GoTo Head1Chk_Cmd_Err
    End If
' * éÛì¸
    If Trim(iH1_Ukeir) <> "" Then
        gSL_Select = " SELECT * "
        gSL_Select = gSL_Select & " FROM éÊà¯êÊéÛì¸É}ÉXÉ^ "
        gSL_Select = gSL_Select & " Where éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & " And   éÊà¯êÊãÊï™ = 1"
        gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
        If Not TUMRead(gSL_Select, 1) Then GoTo Head1Chk_Cmd_Err
    End If
    
'
    gSL_Select = "Select * From çﬁóøî≠íçì‡é¶ÉeÅ[ÉuÉã"
    gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & " And   çﬁéø = '" & RTrim(iH1_Zaist) & "'"
    gSL_Select = gSL_Select & " And   î¬å˙ =  " & CisFun.Val2(iH1_Itatu)
    gSL_Select = gSL_Select & " And    ïù  =  " & CisFun.Val2(iH1_Width)
    gSL_Select = gSL_Select & " And   í∑Ç≥ =  " & CisFun.Val2(iH1_Long)
    gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    If Not DBInput("INV") Then GoTo Head1Chk_Cmd_Err
    
    Head1Chk_Cmd = True
    Exit Function

Head1Chk_Cmd_Err:
    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = "    ÉpÉâÉÅÅ[É^Ç™ê≥ÇµÇ≠ê›íËÇ≥ÇÍÇƒÇ¢Ç»Ç¢Ç©            "
        .MB_MSG(3) = "    Å@Å@Å@Å@ÉfÅ[É^Ç™Ç†ÇËÇ‹ÇπÇÒ            "
        .MB_Title = ""
        .MB_Button = Error
        .MBOX
    End With
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
    
    If Not ZNJRead(gSL_Select, 1) Then
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
        If SyoriKB = "ADD" Then
            gSL_Select = "Select çﬁóøä«óùî‘çÜ, ê°ñ@ãÊï™ From çﬁóøÉ}ÉXÉ^"
            gSL_Select = gSL_Select & " Where çﬁéø = '" & RTrim$(iH1_Zaist) & "'"
            gSL_Select = gSL_Select & " And   î¬å˙ =  " & CisFun.Val2(iH1_Itatu)
            gSL_Select = gSL_Select & " And    ïù  =  " & CisFun.Val2(iH1_Width)
            gSL_Select = gSL_Select & " And   í∑Ç≥ =  " & CisFun.Val2(iH1_Long)
            If ZRMRead(gSL_Select, 1) Then
                hd_ZaiKnrNo = ZRM.çﬁóøä«óùî‘çÜ
                hd_SunpKB = ZRM.ê°ñ@ãÊï™
            End If
        End If
    Else
        iH1_NJYM = RTrim(ZNJ.ì‡é¶îNåé)
        iH1_YM = RTrim(ZNJ.åvâÊîNåé)
        iH1_Zaist = RTrim(ZNJ.çﬁéø)
        iH1_Itatu = Format(ZNJ.î¬å˙, "###0.00;###0.00;#")
        iH1_Width = Format(ZNJ.ïù, "###0.00;###0.00;#")
        iH1_Long = Format(ZNJ.í∑Ç≥, "###0.00;###0.00;#")
        iH1_Torcd = RTrim(ZNJ.éËîzêÊ)
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RTrim(ZNJ.éÛì¸)

        hd_ZaiKnrNo = ZNJ.çﬁóøä«óùî‘çÜ
        hd_SunpKB = ZNJ.ê°ñ@ãÊï™
        
        Call BodySet
        DBInput = True
    End If
    
    Call SetWeek(iH1_YM, iH1_Torcd, iH1_Ukeir)
    Call Body12Set

DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'*-----------------------------------------------------*
'*          É{ÉfÉB(ì˙íˆèÓïÒ)ïîÉZÉbÉg < B1_Area1 >       *
'*-----------------------------------------------------*
Private Sub BodySet()
    Static sw_NJYM          As Date
    Static sw_YM            As Date
    Static sw_MConst        As Byte

    sw_NJYM = Format(RTrim$(iH1_NJYM) & "01", "@@@@/@@/@@")
    sw_YM = Format(RTrim$(iH1_YM) & "01", "@@@@/@@/@@")
    sw_MConst = 50 + DateDiff("M", sw_NJYM, sw_YM)

'*---------*
'* êîó èÓïÒ
'*---------*
'   < ì˙ó êî >
    gSL_Select = ""
    gSL_Select = gSL_Select & " Select '1' KB, SubString( î[ä˙, 7, 2 ) DD, êîó " & vbCrLf
    gSL_Select = gSL_Select & " From çﬁóøî≠íçì‡é¶ïKóvêîëOâÒÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo & vbCrLf
    gSL_Select = gSL_Select & " And   éËîzêÊ   = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸     = '" & RTrim(iH1_Ukeir) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'" & vbCrLf

    gSL_Select = gSL_Select & " Union ALL" & vbCrLf
    gSL_Select = gSL_Select & " Select '2' KB, SubString( î[ä˙, 7, 2 ) DD, êîó " & vbCrLf
    gSL_Select = gSL_Select & " From çﬁóøî≠íçì‡é¶ëOâÒÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo & vbCrLf
    gSL_Select = gSL_Select & " And   éËîzêÊ   = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸     = '" & RTrim(iH1_Ukeir) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'" & vbCrLf

    gSL_Select = gSL_Select & " Union ALL" & vbCrLf
    gSL_Select = gSL_Select & " Select '3' KB, SubString( î[ä˙, 7, 2 ) DD, êîó " & vbCrLf
    gSL_Select = gSL_Select & " From çﬁóøî≠íçì‡é¶ÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo & vbCrLf
    gSL_Select = gSL_Select & " And   éËîzêÊ   = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸     = '" & RTrim(iH1_Ukeir) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'" & vbCrLf

    gSL_Select = gSL_Select & " Union ALL" & vbCrLf
    gSL_Select = gSL_Select & " Select '4' KB, SubString( î[ä˙, 7, 2 ) DD, êîó " & vbCrLf
    gSL_Select = gSL_Select & " From çﬁóøî≠íçì‡é¶ïKóvêîÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo & vbCrLf
    gSL_Select = gSL_Select & " And   éËîzêÊ   = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸     = '" & RTrim(iH1_Ukeir) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'" & vbCrLf
    
'   < åéó êî >
    gSL_Select = gSL_Select & " Union ALL" & vbCrLf
    gSL_Select = gSL_Select & " Select '1' KB,"
    gSL_Select = gSL_Select & " DateDiff( M, ì‡é¶îNåé + '01', åvâÊîNåé + '01' ) + 50 DD,"
    gSL_Select = gSL_Select & " Sum( êîó  ) êîó " & vbCrLf
    gSL_Select = gSL_Select & " From çﬁóøî≠íçì‡é¶ïKóvêîëOâÒÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo & vbCrLf
    gSL_Select = gSL_Select & " And   éËîzêÊ    = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸      = '" & RTrim(iH1_Ukeir) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé >= '" & RTrim(iH1_YM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé  = '" & RTrim(iH1_NJYM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " Group By ì‡é¶îNåé, åvâÊîNåé" & vbCrLf
    
    gSL_Select = gSL_Select & " Union ALL" & vbCrLf
    gSL_Select = gSL_Select & " Select '2' KB,"
    gSL_Select = gSL_Select & " DateDiff( M, ì‡é¶îNåé + '01', åvâÊîNåé + '01' ) + 50 DD,"
    gSL_Select = gSL_Select & " Sum( êîó  ) êîó " & vbCrLf
    gSL_Select = gSL_Select & " From çﬁóøî≠íçì‡é¶ëOâÒÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo & vbCrLf
    gSL_Select = gSL_Select & " And   éËîzêÊ    = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸      = '" & RTrim(iH1_Ukeir) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé >= '" & RTrim(iH1_YM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé  = '" & RTrim(iH1_NJYM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " Group By ì‡é¶îNåé, åvâÊîNåé" & vbCrLf

    gSL_Select = gSL_Select & " Union ALL" & vbCrLf
    gSL_Select = gSL_Select & " Select '3' KB,"
    gSL_Select = gSL_Select & " DateDiff( M, ì‡é¶îNåé + '01', åvâÊîNåé + '01' ) + 50 DD,"
    gSL_Select = gSL_Select & " Sum( êîó  ) êîó " & vbCrLf
    gSL_Select = gSL_Select & " From çﬁóøî≠íçì‡é¶ÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo & vbCrLf
    gSL_Select = gSL_Select & " And   éËîzêÊ    = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸      = '" & RTrim(iH1_Ukeir) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé >= '" & RTrim(iH1_YM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé  = '" & RTrim(iH1_NJYM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " Group By ì‡é¶îNåé, åvâÊîNåé" & vbCrLf

    gSL_Select = gSL_Select & " Union ALL" & vbCrLf
    gSL_Select = gSL_Select & " Select '4' KB,"
    gSL_Select = gSL_Select & " DateDiff( M, ì‡é¶îNåé + '01', åvâÊîNåé + '01' ) + 50 DD,"
    gSL_Select = gSL_Select & " Sum( êîó  ) êîó " & vbCrLf
    gSL_Select = gSL_Select & " From çﬁóøî≠íçì‡é¶ïKóvêîÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo & vbCrLf
    gSL_Select = gSL_Select & " And   éËîzêÊ    = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸      = '" & RTrim(iH1_Ukeir) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé >= '" & RTrim(iH1_YM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé  = '" & RTrim(iH1_NJYM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " Group By ì‡é¶îNåé, åvâÊîNåé" & vbCrLf
    
'   < ïiî‘íPà Å@ìñåéêîó  >
    gSL_Select = gSL_Select & " Union ALL" & vbCrLf
    gSL_Select = gSL_Select & " Select '3' KB, 70 DD, IsNull( Sum( êîó  ), 0 ) êîó " & vbCrLf
    gSL_Select = gSL_Select & " From çﬁóøî≠íçì‡é¶ÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'" & vbCrLf
    
'   < ïiî‘íPà Å@ëOåéï™ìñåéêîó  >
    gSL_Select = gSL_Select & " Union ALL" & vbCrLf
    gSL_Select = gSL_Select & " Select '3' KB, 71 DD, IsNull( Sum( êîó  ), 0 ) êîó " & vbCrLf
    gSL_Select = gSL_Select & " From çﬁóøî≠íçì‡é¶ÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé = Convert( Varchar(6), DateAdd( M, -1, '" & RTrim$(iH1_YM) & "' + '01' ), 112 )" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé = Convert( Varchar(6), DateAdd( M, -1, '" & RTrim$(iH1_NJYM) & "' + '01' ), 112 )" & vbCrLf
    
    gSL_Select = gSL_Select & " Order by KB, DD" & vbCrLf
    If ZNJRead(gSL_Select) Then
        Do Until Not ZNJ_RDSTS
            With ZNJ

                If .DD <= 31 Then
'                   [ ì˙ó  ]
                    Select Case .KB
                        Case "1"
                                    B1lb_LNJSu(.DD) = Format(.êîó , "##0.000;-##0.000;#")   '< ïKóvêî [ ëOâÒ ] >
                        Case "2"
                                    B1lb_LSu(.DD) = Format(.êîó , "##0.000;-##0.000;#")     '< åvâÊêî [ ëOâÒ ] >
                                    B1lb_LSu(.DD).Tag = "X"                                 '  çXêVéûÇ…Ç±ÇÃÉ^ÉOÇégóp
                        Case "3"
                                    iB1_Su(.DD) = CisFun.RSetFld(.êîó , 10, "##0.000;-##0.000;#")   '< åvâÊêî >
                                    iB1_Su(.DD).Tag = .êîó 
                                    If .êîó  <> CisFun.Val2(B1lb_LSu(.DD)) And _
                                       B1lb_LSu(.DD).Tag = "X" Then B1lb_LSu(.DD).BackColor = mHHenColor
                        Case "4"
                                    B1lb_NJSu(.DD) = Format(.êîó , "##0.000;-##0.000;#")    '< ïKóvêî >
                    End Select
                Else
                    If .DD >= sw_MConst And .DD <= sw_MConst + 2 Then
'                       [ åéó  ]
                        Select Case .KB
                            Case "1":   B1lb_LGNJSu(.DD - sw_MConst) = Format(.êîó , "#,##0.000;-#,##0.000;#")      '< ïKóvêî [ ëOâÒ ] >
                            Case "2":   B1lb_LGSu(.DD - sw_MConst) = Format(.êîó , "#,##0.000;-#,##0.000;#")        '< åvâÊêî [ ëOâÒ ] >
                            Case "3"
                                        B1lb_GSu(.DD - sw_MConst) = Format(.êîó , "#,##0.000;-#,##0.000;#")         '< åvâÊêî >
                                        B1lb_GSu(.DD - sw_MConst).Tag = .êîó 
                            Case "4":   B1lb_GNJSu(.DD - sw_MConst) = Format(.êîó , "#,##0.000;-#,##0.000;#")       '< ïKóvêî >
                        End Select
                    Else
'                       [ ïiî‘íPà Å@åéó  ]
                        If .DD = 70 Then B1lb_NJSuT = Format(.êîó , "#,##0.000;-#,##0.000;#")   '< ìñåéïiî‘êîó  >
                        If .DD = 71 Then
                            B1lb_ZNJSuT = Format(.êîó , "#,##0.000;-#,##0.000;#")               '< ëOåéïiî‘êîó  >
                            B1lb_SaNJSuT = Format(CisFun.Val2(B1lb_NJSuT) - .êîó , "#,##0.000;-##,#0.000;#") '< ç∑ >
                        End If
                    End If
                End If
            End With

            Call ZNJReadNext
        Loop
    End If
    Call ZNJClose

'   < ç°âÒåvâÊêîÇ™Ç»Ç≠ÇƒÅAëOâÒåvâÊêîÇ†ÇËÇÃÉoÉbÉNÉJÉâÅ[ê›íË >Å@(è„ãLÇÕÅAç°âÒêîÇ™Ç†Ç¡ÇƒÇÃëOâÒílämîF)
    For gInt = 1 To 31
        If CisFun.Val2(B1lb_LSu(gInt)) <> 0 And CisFun.Val2(iB1_Su(gInt)) = 0 Then
            B1lb_LSu(gInt).BackColor = mHHenColor
        End If
    Next gInt

End Sub
'*------------------------------------------------------*
'*          É{ÉfÉB(ïiî‘èÓïÒ)ïîÉZÉbÉg < B1_Area2 >        *
'*------------------------------------------------------*
Private Sub Body12Set()
    sv_Choku = ""
'*---------*
'* çﬁóøèÓïÒ
'*---------*
    Call ItemsClearZTM
    
    gSL_Select = "Select HT.*," & vbCrLf
    gSL_Select = gSL_Select & " IsNull( HN.ç›å…ä«óùãÊï™, 0 ) ç›å…ä«óùãÊï™," & vbCrLf
    gSL_Select = gSL_Select & " IsNull( HN.ç›å…ä«óùêîíl, 0 ) ç›å…ä«óùêîíl," & vbCrLf
    gSL_Select = gSL_Select & " IsNull( SY.ílñºèÃ, '' ) ç›å…ä«óùñº" & vbCrLf
    gSL_Select = gSL_Select & " From çﬁóøéÊà¯êÊÉ}ÉXÉ^ HT" & vbCrLf
    gSL_Select = gSL_Select & " Left Outer Join çﬁóøÉ}ÉXÉ^ HN" & vbCrLf
    gSL_Select = gSL_Select & " On  HN.çﬁóøä«óùî‘çÜ = HT.çﬁóøä«óùî‘çÜ" & vbCrLf
    gSL_Select = gSL_Select & " Left Outer Join ñºèÃÉ}ÉXÉ^ SY" & vbCrLf
    gSL_Select = gSL_Select & " On  SY.ãÊï™ñºèÃ = 'ç›å…ãÊï™'" & vbCrLf
    gSL_Select = gSL_Select & " And SY.íl = HN.ç›å…ä«óùãÊï™" & vbCrLf
    gSL_Select = gSL_Select & " And SY.ãÊï™É^ÉCÉv = ''" & vbCrLf
    gSL_Select = gSL_Select & " Where HT.çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo & vbCrLf
    gSL_Select = gSL_Select & " And   HT.éÊà¯êÊ = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   HT.éÛì¸   = '" & RTrim(iH1_Ukeir) & "'" & vbCrLf
    If ZTMRead(gSL_Select, 1) Then
        With ZTM
            
'           [ LT ]
            B1lb_LT = Format(.LT, "##")

'           [ ÉçÉbÉg ]
            B1lb_Lot = Format(.ÉçÉbÉg1, "#,###")

'           [ ójì˙éwíË ]
            If .Week1 <> 0 Then: B1lb_Week(1) = "åé"
            If .Week2 <> 0 Then: B1lb_Week(2) = "âŒ"
            If .Week3 <> 0 Then: B1lb_Week(3) = "êÖ"
            If .Week4 <> 0 Then: B1lb_Week(4) = "ñÿ"
            If .Week5 <> 0 Then: B1lb_Week(5) = "ã‡"

'           [ ì˙Ç…ÇøéwíË ]
            If .Day1 <> 0 Then: B1lb_Day(1) = Format(.Day1, "##")
            If .Day2 <> 0 Then: B1lb_Day(2) = Format(.Day2, "##")
            If .Day3 <> 0 Then: B1lb_Day(3) = Format(.Day3, "##")
            If .Day4 <> 0 Then: B1lb_Day(4) = Format(.Day4, "##")
            If .Day5 <> 0 Then: B1lb_Day(5) = Format(.Day5, "##")

'           [ SK ]
            If .SKãÊï™ = 0 Then B1lb_SKKbn = "Ç©ÇÒÇŒÇÒ" Else B1lb_SKKbn = "éwé¶"
            NaiHatu_Back(1) = B1lb_SKKbn
        
'           [ ÉTÉCÉNÉã ]
            B1lb_Cycle = Format(RTrim$(.ÉTÉCÉNÉã), "@-@@-@@")
        
'           [ é˚óeêî ån ]
            B1lb_Syuyo1 = Format(.é˚óeêî1, "#,###")
            B1lb_Syuyo2 = Format(.é˚óeêî2, "#,###")
            If RTrim$(.ïœçXì˙) <> "" Then
                B1lb_CDate = Format(.ïœçXì˙, "@@@@/@@/@@")
                B1lb_CBin = Format(.ïœçXï÷, "##")
            End If
        
'           [ åWêî ]
            B1lb_Kesu1 = Format(.åWêî1, "0.00;0.00;#")
            B1lb_Kesu2 = Format(.åWêî2, "0.00;0.00;#")
            B1lb_Kesu3 = Format(.åWêî3, "0.00;0.00;#")
        
'           [ ç›å…ä«óù ]
            B1lb_ZKKbn = RTrim$(.ç›å…ä«óùñº)
            Select Case .ç›å…ä«óùãÊï™
                Case 2:             B1lb_ZKSu = Format(.ç›å…ä«óùêîíl, "###") & " Åì"
                Case 3:             B1lb_ZKSu = Format(.ç›å…ä«óùêîíl, "###,###")
                Case 4:             B1lb_ZKSu = Format(.ç›å…ä«óùêîíl, "#0.0") & " ì˙"
                Case Else:          B1lb_ZKSu = ""
            End Select
        
'           [ ì‡é¶î≠íç ]
            If .ì‡é¶î≠íçãÊï™ = 0 Then B1lb_NaHaKB = "ÇµÇ»Ç¢" Else B1lb_NaHaKB = "Ç∑ÇÈ"
            NaiHatu_Back(2) = B1lb_NaHaKB
        
'           [ î‰ó¶ ]
            B1lb_Ritu = Format(.î≠íçî‰ó¶, "#Åì;#;#")
    
'           [ óLå¯ä˙å¿ ]
            B1lb_SDate = Format(RTrim$(.óLå¯ä˙å¿äJén), "@@@@/@@/@@")
            B1lb_EDate = Format(RTrim$(.óLå¯ä˙å¿èIóπ), "@@@@/@@/@@")
        
'           [ íºëóêÊ ]
            If .éËîzãÊï™ = 1 Then sv_Choku = RTrim$(.ÉÅÅ[ÉJÅ[)
        
        End With
    End If

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/     ëOÅ@åéÅ@çﬁÅ@óøÅ@ññÅ@ç›Å@å…Å@éZÅ@èo
'_/
    With CisDB
        .SQL = "çﬁóøç›å…êîéZèo"
        .StoadoCount = 5
        .ParaNo = 0:    .ParaIO = Returnóp
        .ParaNo = 1:    .ParaIO = Inputóp
        .ParaNo = 2:    .ParaIO = Inputóp
        .ParaNo = 3:    .ParaIO = Inputóp
        .ParaNo = 4:    .ParaIO = OutPutóp
        .ParaNo = 5:    .ParaIO = OutPutóp
        .ParaNo = 1:    .ParaValue = RTrim$(iH1_YM)
        .ParaNo = 2:    .ParaValue = hd_ZaiKnrNo
        .ParaNo = 3:    .ParaValue = 1
        .DBStored
    
        .ParaNo = 0
        If .ParaValue <> 0 Then
            CisFun.MB_Lines = 5
            CisFun.MB_MSG(2) = "    ëOåéññç›å…ÇÃéZèoÇ…é∏îsÇµÇ‹ÇµÇΩ    "
            CisFun.MB_Button = OK
            CisFun.MBOX
            Exit Sub
        End If
    
        .ParaNo = 5:    B1lb_Zaiko = Format(.ParaValue, "#,###")
    End With
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/     SK = 1(éwé¶)Ç≈ÅAì‡é¶î≠íç = 1(Ç∑ÇÈ)ÇÃèÍçáÅAåªç›î≠íççœêîÇéZèoÇµï\é¶Ç∑ÇÈ
'_/
    Static sw_Flg           As Boolean
    
    If ZTM.SKãÊï™ = 1 And ZTM.ì‡é¶î≠íçãÊï™ = 1 Then

        gSL_Select = "Select Sum( î≠íçêî ) î≠íçêî"
        gSL_Select = gSL_Select & " From çﬁóøî≠íçÉeÅ[ÉuÉã"
        gSL_Select = gSL_Select & " Where çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo
        gSL_Select = gSL_Select & " And   édì¸êÊ = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & " And   éÛì¸   = '" & RTrim(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & " And   SubString( î[ì¸ì˙, 1, 6 ) = '" & RTrim$(iH1_YM) & "'"
        If ZCTRead(gSL_Select, 1) Then B1lb_HCSu = Format(ZCT.î≠íçêî, "#,##0.000;-#,##0.000;#")
        sw_Flg = True
    Else
        sw_Flg = False
    End If

    NaiHatu_Back(0).Visible = sw_Flg
    NaiHatu_Back(1).Visible = sw_Flg
    NaiHatu_Back(2).Visible = sw_Flg
    B1lb_HCSu.Visible = sw_Flg
    B1lb_SKKbn.Visible = Not sw_Flg
    B1lb_NaHaKB.Visible = Not sw_Flg

'_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
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
    If Not CheckMessage Then Exit Function
    
    With CisDB
        On Error GoTo DBDelete_Err
        .DBTran (TransBegin) ' ƒ◊›ªﬁ∏ºÆ›äJén
' ( çÌèúèàóù )
        .SQL = "DELETE  çﬁóøî≠íçì‡é¶ÉeÅ[ÉuÉã "
        .SQL = .SQL & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        .SQL = .SQL & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
        .SQL = .SQL & " And   çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo
        .SQL = .SQL & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
        .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
        .DBExec
        
        .SQL = "DELETE  çﬁóøî≠íçì‡é¶ëOâÒÉeÅ[ÉuÉã "
        .SQL = .SQL & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        .SQL = .SQL & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
        .SQL = .SQL & " And   çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo
        .SQL = .SQL & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
        .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
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
    Dim AddFlg          As Boolean
    Static ClwFlg       As Boolean
    Static wsSYMD       As String
    Static wsTYMD       As String
' *ADD*
    Static MntFlg       As Boolean

'  ( ämîFÉÅÉbÉZÅ[ÉW )
    If Not CheckMessage Then Exit Function
    
    On Error GoTo DBPut_Err
    
'( ÉÅÉìÉeópÉJÉåÉìÉ_ÉèÅ[ÉNçÏê¨ )
    With CisDB
        .SQL = "ÉÅÉìÉeópÉJÉåÉìÉ_ÉèÅ[ÉNçÏê¨"
        .StoadoCount = 3
        .ParaNo = 0:    .ParaIO = Returnóp
        .ParaNo = 1:    .ParaIO = Inputóp
        .ParaNo = 2:    .ParaIO = Inputóp
        .ParaNo = 3:    .ParaIO = Inputóp
        .ParaNo = 1:    .ParaValue = RTrim(iH1_YM)
        .ParaNo = 2:    .ParaValue = RTrim(iH1_Torcd)
        .ParaNo = 3:    .ParaValue = RTrim(iH1_Ukeir)
        .DBStored
    
        .ParaNo = 0
        If .ParaValue <> 0 Then
            CisFun.MB_Lines = 5
            CisFun.MB_MSG(2) = "    ÉJÉåÉìÉ_ÉèÅ[ÉNçÏê¨Ç…é∏îsÇµÇ‹ÇµÇΩÅB    "
            CisFun.MB_Button = OK
            CisFun.MBOX
            Exit Function
        End If
    End With
    
' *ADD* 2004/04/09
    MntFlg = False
    For gInt = 1 To 31
        If CisFun.Val2(iB1_Su(gInt)) <> CisFun.Val2(iB1_Su(gInt).Tag) Then
            MntFlg = True
            Exit For
        End If
    Next gInt
    
    If MntFlg And SyoriKB = "MNT" Then
        With CisDB
        ' *ëOâÒì‡óe çÌèú*
            .SQL = "DELETE çﬁóøî≠íçì‡é¶ëOâÒÉeÅ[ÉuÉã "
            .SQL = .SQL & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
            .SQL = .SQL & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
            .SQL = .SQL & " And   çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo
            .SQL = .SQL & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
            .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
            .DBExec
            
        ' *ëOâÒÉfÅ[É^ çÏê¨*
            .SQL = "INSERT INTO çﬁóøî≠íçì‡é¶ëOâÒÉeÅ[ÉuÉã "
            .SQL = .SQL & " ( ì‡é¶îNåé,çﬁóøä«óùî‘çÜ,éËîzêÊ,éÛì¸,íºëóêÊ, "
            .SQL = .SQL & "   î[ä˙,édä|ì˙,éËîzì˙,çﬁéø,ê°ñ@ãÊï™,î¬å˙,ïù,í∑Ç≥, "
            .SQL = .SQL & "   êîó ,èàóùãÊï™,çÌèúãÊï™,ïœçXâÒêî,åvâÊîNåé,"
            .SQL = .SQL & "   çÏê¨é“,çÏê¨ì˙,çÏê¨í[ññ )"
            .SQL = .SQL & " SELECT ì‡é¶îNåé,çﬁóøä«óùî‘çÜ,éËîzêÊ,éÛì¸,íºëóêÊ,"
            .SQL = .SQL & "        î[ä˙,édä|ì˙,éËîzì˙,çﬁéø,ê°ñ@ãÊï™,î¬å˙,ïù,í∑Ç≥, "
            .SQL = .SQL & "        êîó ,èàóùãÊï™,çÌèúãÊï™,ïœçXâÒêî,åvâÊîNåé,"
            .SQL = .SQL & "'" & gTanto & "',"
            .SQL = .SQL & "GETDATE(),HOST_NAME() "
            .SQL = .SQL & "        FROM çﬁóøî≠íçì‡é¶ÉeÅ[ÉuÉã  "
            .SQL = .SQL & "        Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
            .SQL = .SQL & "          And åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
            .SQL = .SQL & "          And çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo
            .SQL = .SQL & "          And éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
            .SQL = .SQL & "          And éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
            .DBExec
        End With
    End If
    
    ClwFlg = False
    For gInt = 1 To 31
        If CisFun.Val2(iB1_Su(gInt)) <> CisFun.Val2(iB1_Su(gInt).Tag) Then
            
            gStr = RTrim(iH1_YM) & Format(gInt, "00")
            gSL_Select = "SELECT * FROM çﬁóøî≠íçì‡é¶ÉeÅ[ÉuÉã "
            gSL_Select = gSL_Select & " WHERE ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
            gSL_Select = gSL_Select & "   AND çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo
            gSL_Select = gSL_Select & "   AND éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
            gSL_Select = gSL_Select & "   AND éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
            gSL_Select = gSL_Select & "   AND î[ä˙ = '" & RTrim(gStr) & "'"
            If ZNJRead(gSL_Select, 1) Then
                AddFlg = False
' *DEL* 2004/04/09
'                With ZNJ
'                    ZZN.ì‡é¶îNåé = .ì‡é¶îNåé
'                    ZZN.çﬁóøä«óùî‘çÜ = .çﬁóøä«óùî‘çÜ
'                    ZZN.éËîzêÊ = .éËîzêÊ
'                    ZZN.éÛì¸ = .éÛì¸
'                    ZZN.íºëóêÊ = .íºëóêÊ
'                    ZZN.î[ä˙ = .î[ä˙
'                    ZZN.édä|ì˙ = .édä|ì˙
'                    ZZN.éËîzì˙ = .éËîzì˙
'                    ZZN.çﬁéø = .çﬁéø
'                    ZZN.ê°ñ@ãÊï™ = .ê°ñ@ãÊï™
'                    ZZN.î¬å˙ = .î¬å˙
'                    ZZN.ïù = .ïù
'                    ZZN.í∑Ç≥ = .í∑Ç≥
'                    ZZN.êîó  = .êîó 
'                    ZZN.èàóùãÊï™ = .èàóùãÊï™
'                    ZZN.çÌèúãÊï™ = .çÌèúãÊï™
'                    ZZN.ïœçXâÒêî = 1
'                    ZZN.åvâÊîNåé = .åvâÊîNåé
'                End With
'
'                If RTrim(B1lb_LSu(gInt).Tag) = "" Then
'                    ZZN.çÏê¨é“ = gTanto
'                    Call ZZNInsert
'                Else
'                    ZZN.çXêVé“ = gTanto
'                    Call ZZNUpdate
'                End If
            Else
                
            ' édä|ÅEéËîzì˙
                With CisDB
                    .SQL = "ÉÅÉìÉeçﬁóøópédä|éËîzì˙éZèo"
                    .StoadoCount = 6
                    .ParaNo = 0:    .ParaIO = Returnóp
                    .ParaNo = 1:    .ParaIO = Inputóp
                    .ParaNo = 2:    .ParaIO = Inputóp
                    .ParaNo = 3:    .ParaIO = Inputóp
                    .ParaNo = 4:    .ParaIO = Inputóp
                    .ParaNo = 5:    .ParaIO = OutPutóp
                    .ParaNo = 6:    .ParaIO = OutPutóp

                    .ParaNo = 1:    .ParaValue = hd_ZaiKnrNo
                    .ParaNo = 2:    .ParaValue = RTrim(iH1_Torcd)
                    .ParaNo = 3:    .ParaValue = RTrim(iH1_Ukeir)
                    .ParaNo = 4:    .ParaValue = RTrim(gStr)

                    .DBStored

                    .ParaNo = 0
                    If .ParaValue <> 0 Then
                        CisFun.MB_Lines = 5
'                        CisFun.MB_MSG(2) = "    ÉJÉåÉìÉ_ÉèÅ[ÉNçÏê¨Ç…é∏îsÇµÇ‹ÇµÇΩÅB    "                'CHG-DEL 2010/02/23
                        CisFun.MB_MSG(2) = "    ÉÅÉìÉeçﬁóøópédä|éËîzì˙éZèoÇ…é∏îsÇµÇ‹ÇµÇΩÅB    "         'CHG-ADD 2010/02/23
                        CisFun.MB_Button = OK
                        CisFun.MBOX
                        Exit Function
                    End If

                    .ParaNo = 5:    wsSYMD = .ParaValue
                    .ParaNo = 6:    wsTYMD = .ParaValue
                
                    On Error Resume Next: .Recordset.Close: On Error GoTo 0
                End With
                
                AddFlg = True
                Call ItemsClearZNJ
                With ZNJ
                    .ì‡é¶îNåé = RTrim(iH1_NJYM)
                    .çﬁóøä«óùî‘çÜ = hd_ZaiKnrNo
                    .éËîzêÊ = RTrim(iH1_Torcd)
                    .éÛì¸ = RTrim(iH1_Ukeir)
                    .íºëóêÊ = RTrim$(sv_Choku)
                    .î[ä˙ = RTrim(gStr)
                    .édä|ì˙ = RTrim(wsSYMD)
                    .éËîzì˙ = RTrim(wsTYMD)
                    .çﬁéø = RTrim(iH1_Zaist)
                    .ê°ñ@ãÊï™ = hd_SunpKB
                    .î¬å˙ = CisFun.Val2(iH1_Itatu)
                    .ïù = CisFun.Val2(iH1_Width)
                    .í∑Ç≥ = CisFun.Val2(iH1_Long)
                End With
            End If
    
            ZNJ.èàóùãÊï™ = 0
            ZNJ.êîó  = CisFun.Val2(iB1_Su(gInt))
            ZNJ.åvâÊîNåé = RTrim(iH1_YM)
            If ZNJ.êîó  <> 0 Then
                If AddFlg Then
                    ZNJ.çÏê¨é“ = gTanto
                    Call ZNJInsert
                    iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
                Else
                    ZNJ.çXêVé“ = gTanto
                    Call ZNJUpdate
                    iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
                End If
            Else
        ' ( çÌèúèàóù )
                With CisDB
                    .SQL = "DELETE çﬁóøî≠íçì‡é¶ÉeÅ[ÉuÉã "
                    .SQL = .SQL & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
                    .SQL = .SQL & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
                    .SQL = .SQL & " And   çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo
                    .SQL = .SQL & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
                    .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
                    .SQL = .SQL & " And   î[ä˙ = '" & RTrim(gStr) & "'"
                    .DBExec
                End With
                iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")
            End If
        
            ClwFlg = True
        End If
    Next gInt
    
'[ ì‡é¶î≠íçèàóùÅ@òAågópÉgÉäÉKÅ[ÉeÅ[ÉuÉãçXêV (ì‡é¶î≠íçÉ`ÉFÉbÉNÉeÅ[ÉuÉã) ]
    If ClwFlg Then
        gSL_Select = "Select ïiî‘ From ì‡é¶î≠íçÉ`ÉFÉbÉNÉeÅ[ÉuÉã"
        gSL_Select = gSL_Select & " Where ãÊï™ = 1"
        gSL_Select = gSL_Select & " And   î≠íçîNåé = '" & RTrim$(iH1_YM) & "'"
        gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim$(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim$(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & " And   çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo
        If NHCRead(gSL_Select, 1) Then
            With CisDB
                .SQL = "Delete ì‡é¶î≠íçÉ`ÉFÉbÉNÉeÅ[ÉuÉã "
                .SQL = .SQL & " Where ãÊï™ = 1"
                .SQL = .SQL & " And   î≠íçîNåé = '" & RTrim$(iH1_YM) & "'"
                .SQL = .SQL & " And   éËîzêÊ = '" & RTrim$(iH1_Torcd) & "'"
                .SQL = .SQL & " And   éÛì¸ = '" & RTrim$(iH1_Ukeir) & "'"
                .SQL = .SQL & " And   çﬁóøä«óùî‘çÜ = " & hd_ZaiKnrNo
                .DBExec
            End With
        End If
    End If
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'*-------------------------------------------------*
'*      Éf Å[ É^ çX êV  äm îF ÉÅ Éb ÉZ Å[ ÉW
'*-------------------------------------------------*
Private Function CheckMessage() As Boolean
    CheckMessage = False
    
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     " & SyoriNM & "èàóùÇçsÇ¢Ç‹Ç∑      "
        .MB_MSG(3) = "     ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH      "
        .MB_Title = SyoriNM & "èàóù"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    CheckMessage = True
End Function
'+---------------------------------------+
'+      ójì˙ÉZÉbÉg
'+---------------------------------------+
Private Function SetWeek(YM As String, Torcd As String, Ukeir As String)
    SetWeek = False

    gSL_Select = "Select éÊà¯êÊ, éÛì¸, â“ìÆãÊï™, ñÈãŒãÊï™, ññì˙" & vbCrLf
    gSL_Select = gSL_Select & " From ÉJÉåÉìÉ_É}ÉXÉ^" & vbCrLf
    gSL_Select = gSL_Select & " Where éÊà¯êÊãÊï™ = 1" & vbCrLf
    gSL_Select = gSL_Select & " And   îN = '" & Mid(YM, 1, 4) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åé = '" & Mid(YM, 5, 2) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÊà¯êÊ = '" & RTrim(Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(Ukeir) & "'" & vbCrLf
    
    gSL_Select = gSL_Select & " Union" & vbCrLf
    gSL_Select = gSL_Select & " Select éÊà¯êÊ, éÛì¸, â“ìÆãÊï™, ñÈãŒãÊï™, ññì˙" & vbCrLf
    gSL_Select = gSL_Select & " From ÉJÉåÉìÉ_É}ÉXÉ^" & vbCrLf
    gSL_Select = gSL_Select & " Where éÊà¯êÊãÊï™ = 1" & vbCrLf
    gSL_Select = gSL_Select & " And   îN = '" & Mid(YM, 1, 4) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åé = '" & Mid(YM, 5, 2) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÊà¯êÊ = '" & RTrim(Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸ = ''" & vbCrLf
    
    gSL_Select = gSL_Select & " Union" & vbCrLf
    gSL_Select = gSL_Select & " Select éÊà¯êÊ, éÛì¸, â“ìÆãÊï™, ñÈãŒãÊï™, ññì˙" & vbCrLf
    gSL_Select = gSL_Select & " From ÉJÉåÉìÉ_É}ÉXÉ^" & vbCrLf
    gSL_Select = gSL_Select & " Where éÊà¯êÊãÊï™ = 1" & vbCrLf
    gSL_Select = gSL_Select & " And   îN = '" & Mid(YM, 1, 4) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åé = '" & Mid(YM, 5, 2) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÊà¯êÊ = ''" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸ = ''" & vbCrLf
    
    gSL_Select = gSL_Select & " Order By éÊà¯êÊ Desc, éÛì¸ Desc" & vbCrLf
    
    If Not CLMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "       Åy ÉJÉåÉìÉ_É}ÉXÉ^ ñ¢ìoò^ Åz      "
            .MB_MSG(4) = "     ÉJÉåÉìÉ_É}ÉXÉ^ÇÃìoò^ÇçsÇ¡Çƒâ∫Ç≥Ç¢      "
            .MB_Title = ""
            .MB_Button = OK
            .MBOX
        End With
        Exit Function
    End If

' â“ì≠èÛãµÉZÉbÉg
    For gInt = 1 To 31
        If gInt <= CLM.ññì˙ Then
            B_Sur(gInt).Visible = True
            
            Select Case CisFun.Val2(Mid(CLM.â“ìÆãÊï™, gInt, 1)) _
                      + CisFun.Val2(Mid(CLM.ñÈãŒãÊï™, gInt, 1))
            
                Case 0
'                       < ëSâ“ìÆ >
                        B1lb_Kad(gInt).Visible = False
                        BLB_DD(gInt).BackColor = mOndayColor
                Case 1
'                       < îºâ“ìÆ >
                        B1lb_Kad(gInt).Visible = True
                        BLB_DD(gInt).BackColor = mOndayColor
                        If Mid(CLM.â“ìÆãÊï™, gInt, 1) = "1" Then B1lb_Kad(gInt) = "ñÈ" Else B1lb_Kad(gInt) = "íã"
                Case Else
'                       < ãxì˙ >
                        B1lb_Kad(gInt).Visible = False
                        BLB_DD(gInt).BackColor = mOffdayColor
            End Select
            
        Else
            B_Sur(gInt).Visible = False
        End If
    Next gInt

' (ìñåé)
    B1lb_NYM(0) = Mid(YM, 1, 4) & "îN" & Mid(YM, 5, 2) & "åé"
' (óÇåé)
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add YM, M, 1
        gStr = .Date_Result
    End With
    B1lb_NYM(1) = Mid(gStr, 1, 4) & "îN" & Mid(gStr, 5, 2) & "åé"
' (óÇÅXåé)
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add YM, M, 2
        gStr = .Date_Result
    End With
    B1lb_NYM(2) = Mid(gStr, 1, 4) & "îN" & Mid(gStr, 5, 2) & "åé"

    SetWeek = True
End Function
'+---------------------------------------+
'+      éËîzêÊÉZÉbÉg
'+---------------------------------------+
Private Function SetTorcd()
    SetTorcd = False

    If SyoriKB = "ADD" Then
        If RTrim$(iH1_Zaist) = "" Then Exit Function

'       < í«â¡éûÅFçﬁóøéÊà¯êÊÉ}ÉXÉ^ >
        gSL_Select = "Select éÊà¯êÊ, éÛì¸"
        gSL_Select = gSL_Select & " From çﬁóøéÊà¯êÊÉ}ÉXÉ^"
        gSL_Select = gSL_Select & " Where çﬁéø = '" & RTrim$(iH1_Zaist) & "'"
        gSL_Select = gSL_Select & " And   î¬å˙ =  " & CisFun.Val2(iH1_Itatu)
        gSL_Select = gSL_Select & " And    ïù  =  " & CisFun.Val2(iH1_Width)
        gSL_Select = gSL_Select & " And   í∑Ç≥ =  " & CisFun.Val2(iH1_Long)
        If ZTMRead(gSL_Select) Then
            If CisDB.RecordCount > 1 Then
                Call ZTMClose
                Call Look_Tehacd
                sv_Zaist = RTrim$(iH1_Zaist)
                sv_Itatu = CisFun.Val2(iH1_Itatu)
                sv_Width = CisFun.Val2(iH1_Width)
                sv_Long = CisFun.Val2(iH1_Long)
            Else
                iH1_Torcd = ZTM.éÊà¯êÊ
                iH1_Ukeir = ZTM.éÛì¸
                Call ZTMClose
            End If
        End If
    
    Else
        If RTrim$(iH1_NJYM) = "" Or _
           RTrim$(iH1_YM) = "" Or _
           RTrim$(iH1_Zaist) = "" Then Exit Function

'       * çﬁóøî≠íçì‡é¶ÉeÅ[ÉuÉã ë∂ç›É`ÉFÉbÉN
        gSL_Select = "SELECT éËîzêÊ, éÛì¸"
        gSL_Select = gSL_Select & " FROM çﬁóøî≠íçì‡é¶ÉeÅ[ÉuÉã "
        gSL_Select = gSL_Select & " WHERE ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        gSL_Select = gSL_Select & "   AND åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
        gSL_Select = gSL_Select & "   AND çﬁéø = '" & RTrim(iH1_Zaist) & "'"
        gSL_Select = gSL_Select & " GROUP BY éËîzêÊ, éÛì¸"
        If ZNJRead(gSL_Select) Then
        ' * éËîzêÊëIë
            If CisDB.RecordCount > 1 Then
                Call ZNJClose
                Call Look_Tehacd
                sv_Zaist = RTrim$(iH1_Zaist)
                sv_Itatu = CisFun.Val2(iH1_Itatu)
                sv_Width = CisFun.Val2(iH1_Width)
                sv_Long = CisFun.Val2(iH1_Long)
            Else
                iH1_Torcd = ZNJ.éËîzêÊ
                iH1_Ukeir = ZNJ.éÛì¸
                Call ZNJClose
            End If
        End If
    End If

SetTorcd_ED:
    SetTorcd = True
End Function
