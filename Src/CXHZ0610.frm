VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{8F4FF474-3E18-11D4-AEC2-00000EA57F9E}#1.0#0"; "CISMSG~1.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXHZ0610 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "çﬁóøÅEéxããé¿ê—èÛãµè∆âÔ"
   ClientHeight    =   11010
   ClientLeft      =   1485
   ClientTop       =   2295
   ClientWidth     =   15015
   BeginProperty Font 
      Name            =   "ÇlÇr ÉSÉVÉbÉN"
      Size            =   11.25
      Charset         =   128
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z µ∞¿ﬁ∞
   ScaleHeight     =   11010
   ScaleWidth      =   15015
   WindowState     =   2  'ç≈ëÂâª
   Begin Cis3D_v60.CIS3D B2_Area1 
      Height          =   3360
      Left            =   7365
      Top             =   3510
      Visible         =   0   'False
      Width           =   7405
      _ExtentX        =   13070
      _ExtentY        =   5927
      BackColor       =   12632319
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cBoderWidth     =   3
      Begin CisMsg_v60.CISMSG CISMSG3 
         Height          =   3255
         Left            =   60
         TabIndex        =   24
         Top             =   60
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   5741
         BackColor       =   -2147483633
         ForeColor       =   65280
         Caption         =   " Å¶çﬁóøÅEéxããé¿ê—ñæç◊Å¶"
         BackStyle       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   8421504
         cBoderColor2    =   -2147483643
         cFont3DColor1   =   16777152
         cFont3DColor2   =   0
         cAlingnment     =   6
         cBoderStyle     =   1
         cPositionY      =   80
         gGC2_Green      =   160
         gGradation      =   -1  'True
      End
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   300
         Index           =   0
         Left            =   2640
         Top             =   60
         Width           =   4730
         _ExtentX        =   8334
         _ExtentY        =   529
         ForeColor       =   16711680
         Caption         =   "çﬁéø"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
            Size            =   12
            Charset         =   128
            Weight          =   700
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
         cPositionX      =   30
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   0
         Left            =   3015
         Top             =   1995
         Width           =   1930
         _ExtentX        =   3413
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "YYYY/MM/DD-99"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
         cAlingnment     =   7
         cBoderStyle     =   1
      End
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   300
         Index           =   1
         Left            =   435
         Top             =   60
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         ForeColor       =   16711680
         Caption         =   "î[ î‘"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
            Size            =   12
            Charset         =   128
            Weight          =   700
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
         cPositionX      =   30
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   315
         Index           =   1
         Left            =   420
         Top             =   360
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         BackColor       =   14737632
         Caption         =   "XXXXX"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   7
         cBoderStyle     =   1
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   315
         Index           =   2
         Left            =   2640
         Top             =   360
         Width           =   4730
         _ExtentX        =   8334
         _ExtentY        =   556
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cBoderStyle     =   1
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   11
         Left            =   435
         Top             =   1005
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "çﬁóøñº"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   21
         Left            =   1455
         Top             =   1005
         Width           =   5915
         _ExtentX        =   10425
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   300
         Index           =   2
         Left            =   1545
         Top             =   60
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         ForeColor       =   16711680
         Caption         =   "îwî‘çÜ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
            Size            =   12
            Charset         =   128
            Weight          =   700
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
         cPositionX      =   30
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   315
         Index           =   3
         Left            =   1545
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         BackColor       =   14737632
         Caption         =   "XXXXXXXX"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cBoderStyle     =   1
      End
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   4
         Left            =   435
         Top             =   2325
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "éxããêÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   9
         Left            =   1485
         Top             =   2325
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   5
         Left            =   5685
         Top             =   2325
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "éÛì¸"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
         cPositionX      =   10
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   10
         Left            =   6330
         Top             =   2325
         Width           =   1040
         _ExtentX        =   1826
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "XX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   6
         Left            =   435
         Top             =   1335
         Width           =   1375
         _ExtentX        =   2434
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "é˚óeêî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   4
         Left            =   435
         Top             =   1665
         Width           =   1375
         _ExtentX        =   2434
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "ZZZ,ZZ0.000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
         cPositionX      =   -30
      End
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   7
         Left            =   435
         Top             =   1995
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "ñáêî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   5
         Left            =   1185
         Top             =   1995
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "ZZ,ZZZ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
         cPositionX      =   -30
      End
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   8
         Left            =   3055
         Top             =   1335
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "é¿ê—êî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   6
         Left            =   3055
         Top             =   1665
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "ZZZ,ZZ0.000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
         cPositionX      =   -30
      End
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   19
         Left            =   4740
         Top             =   2985
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "íuèÍ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   15
         Left            =   5685
         Top             =   2985
         Width           =   1675
         _ExtentX        =   2963
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   20
         Left            =   2730
         Top             =   2985
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "ÉTÉCÉNÉã"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   16
         Left            =   3765
         Top             =   2985
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "XXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
         cAlingnment     =   7
         cBoderStyle     =   1
      End
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   3
         Left            =   1845
         Top             =   1995
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "é¿ê—ì˙-ï÷"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
         cPositionX      =   30
      End
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   25
         Left            =   5815
         Top             =   1335
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "ã‡ äz"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   20
         Left            =   5815
         Top             =   1665
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "ZZ,ZZZ,ZZZ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
         cPositionX      =   -30
      End
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   9
         Left            =   435
         Top             =   2985
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "√ﬁ∞¿ãÊï™"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   12
         Left            =   1485
         Top             =   2985
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "XXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
         cAlingnment     =   7
         cBoderStyle     =   1
      End
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   10
         Left            =   435
         Top             =   675
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "ê°ñ@"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   8
         Left            =   1455
         Top             =   675
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "9999.99Å~9999.99Å~9999.99"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   12
         Left            =   5670
         Top             =   675
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "íPà "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   7
         Left            =   6270
         Top             =   675
         Width           =   1100
         _ExtentX        =   1931
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "XXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   14
         Left            =   4450
         Top             =   1335
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "é¿ê—ä∑éZíl"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   18
         Left            =   4450
         Top             =   1665
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "ZZZ,ZZ0.000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
         cPositionX      =   -30
      End
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   27
         Left            =   4945
         Top             =   1995
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "çﬁéÌ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   13
         Left            =   5440
         Top             =   1995
         Width           =   1930
         _ExtentX        =   3413
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXX5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   28
         Left            =   435
         Top             =   2655
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "å` èÛ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   14
         Left            =   1485
         Top             =   2655
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXX5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   13
         Left            =   3330
         Top             =   2655
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "çﬁéÌ∫∞ƒﬁ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   11
         Left            =   4350
         Top             =   2655
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXX5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   15
         Left            =   6180
         Top             =   2655
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "óLñ≥"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
         cPositionX      =   10
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   17
         Left            =   6690
         Top             =   2655
         Width           =   680
         _ExtentX        =   1191
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "XX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   330
         Index           =   16
         Left            =   1810
         Top             =   1335
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "íPâø"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   19
         Left            =   1810
         Top             =   1665
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "ZZZ,ZZ0.00"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
         cPositionX      =   -30
      End
   End
   Begin Cis3D_v60.CIS3D PNL_Wait 
      Height          =   930
      Left            =   5100
      Top             =   3870
      Visible         =   0   'False
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   1640
      BackColor       =   16761024
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
         Caption         =   "ÅyíäèoíÜÅzÇµÇŒÇÁÇ≠Ç®ë“Çøâ∫Ç≥Ç¢"
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
         cBoderWidth     =   2
         cAlingnment     =   7
         cBoderStyle     =   1
      End
   End
   Begin VB.PictureBox PNL_ExeMsg 
      Height          =   2325
      Left            =   500
      ScaleHeight     =   2265
      ScaleWidth      =   6570
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5925
      Visible         =   0   'False
      Width           =   6630
      Begin Cis3D_v60.CIS3D lb_SKensu 
         Height          =   345
         Left            =   3930
         Top             =   480
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   609
         ForeColor       =   0
         Caption         =   "Åy èWåvåèêî ÅÅ ZZZZZÅz"
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
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   780
         Left            =   585
         Top             =   1320
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   1376
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   8421504
         cBoderColor2    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderStyle     =   1
         Begin Cis3D_v60.CIS3D FXLabel7 
            Height          =   345
            Left            =   45
            Top             =   45
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   609
            ForeColor       =   0
            Caption         =   "é¿ê—êî"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D lb_THSu 
            Height          =   345
            Left            =   45
            Top             =   390
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   609
            BackColor       =   16777215
            ForeColor       =   0
            Caption         =   "ZZZ,ZZ0.000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cPositionX      =   -60
         End
         Begin Cis3D_v60.CIS3D lb_TNSu 
            Height          =   345
            Left            =   1770
            Top             =   390
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   609
            BackColor       =   16777215
            ForeColor       =   0
            Caption         =   "ZZZ,ZZZ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cPositionX      =   -60
         End
         Begin Cis3D_v60.CIS3D FXLabel8 
            Height          =   345
            Left            =   1785
            Top             =   45
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   609
            ForeColor       =   0
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D lb_TNYSu 
            Height          =   345
            Left            =   3525
            Top             =   390
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   609
            BackColor       =   16777215
            ForeColor       =   0
            Caption         =   "ZZZ,ZZZ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cPositionX      =   -60
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   345
            Left            =   3525
            Top             =   45
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   609
            ForeColor       =   0
            Caption         =   "é¿ê—ã‡äz"
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
            cAlingnment     =   7
         End
      End
      Begin Cis3D_v60.CIS3D lb_TMei 
         Height          =   465
         Left            =   45
         Top             =   825
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   820
         ForeColor       =   16711680
         Caption         =   "çﬁéøÅFXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   200
      End
      Begin Cis3D_v60.CIS3D PNL_ExeMsgX 
         Height          =   345
         Left            =   45
         Top             =   480
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   609
         ForeColor       =   0
         Caption         =   "Åy èWåvíPà  ÅÅ é¿ê—ì˙ - ï÷ Åz"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D FXLabel3 
         Height          =   1830
         Left            =   -30
         Top             =   420
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   3228
         ForeColor       =   16576
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
      End
      Begin CisMsg_v60.CISMSG CISMSG2 
         Height          =   570
         Left            =   0
         TabIndex        =   23
         Top             =   -90
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ForeColor       =   12583104
         Caption         =   "Å¶  èW åv åã â   Å¶"
         BackStyle       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   8421504
         cBoderColor2    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   0
         cAlingnment     =   7
         cBoderStyle     =   1
         gGC2_Green      =   160
         gGradation      =   -1  'True
      End
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   7890
      Left            =   285
      ScaleHeight     =   7830
      ScaleWidth      =   14535
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2730
      Width           =   14595
      Begin TrueOleDBGrid60.TDBGrid DBGrid1 
         Bindings        =   "CXHZ0610.frx":0000
         Height          =   7815
         Left            =   60
         OleObjectBlob   =   "CXHZ0610.frx":0015
         TabIndex        =   34
         Top             =   75
         Width           =   14535
      End
   End
   Begin VB.PictureBox H1_Area1 
      Height          =   1275
      Left            =   270
      ScaleHeight     =   1215
      ScaleWidth      =   12990
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1260
      Width           =   13050
      Begin Cis3D_v60.CIS3D UC_3D9 
         Height          =   1125
         Index           =   2
         Left            =   10155
         Top             =   45
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1984
         ForeColor       =   16711680
         Caption         =   "î[Å@î‘"
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
         Begin CisText_V60.CisText iH1_SNouNo 
            Height          =   360
            Left            =   450
            TabIndex        =   6
            Top             =   360
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   635
            cWildeCard      =   -1  'True
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
         Begin CisText_V60.CisText iH1_ENouNo 
            Height          =   360
            Left            =   450
            TabIndex        =   7
            Top             =   720
            Width           =   735
            _ExtentX        =   1296
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
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   225
            Left            =   30
            Top             =   405
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   397
            Caption         =   "äJén"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cBoderWidth     =   2
            cFont3D         =   2
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   255
            Left            =   45
            Top             =   765
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   450
            Caption         =   "èIóπ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cBoderWidth     =   2
            cFont3D         =   2
            cBoderStyle     =   2
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   570
         Left            =   2925
         Top             =   45
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   1005
         ForeColor       =   16711680
         Caption         =   "éxããêÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         cPositionX      =   70
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   960
            TabIndex        =   4
            Top             =   135
            Width           =   990
            _ExtentX        =   1746
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
         Begin Cis3D_v60.CIS3D H1lb_Torcd 
            Height          =   360
            Left            =   1950
            Top             =   135
            Width           =   2460
            _ExtentX        =   4339
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
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   570
         Left            =   7485
         Top             =   45
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1005
         ForeColor       =   16711680
         Caption         =   "éÛì¸"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         cPositionX      =   100
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   360
            Left            =   600
            TabIndex        =   5
            Top             =   105
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
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   1125
         Left            =   11400
         Top             =   45
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   1984
         ForeColor       =   16711680
         Caption         =   "ÉfÅ[É^ãÊï™"
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
         Begin VB.OptionButton OpH1_DspKbn 
            BackColor       =   &H00C0FFC0&
            Caption         =   "ëSÇƒ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   28
            Top             =   315
            Width           =   1050
         End
         Begin VB.OptionButton OpH1_DspKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "éxãã"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   135
            TabIndex        =   27
            Top             =   585
            Width           =   1050
         End
         Begin VB.OptionButton OpH1_DspKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "íºëóéxãã"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   135
            TabIndex        =   26
            Top             =   855
            Width           =   1290
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   1125
         Index           =   0
         Left            =   45
         Top             =   45
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   1984
         ForeColor       =   16711680
         Caption         =   "é¿ ê— ì˙-ï÷"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisYMD_With_Btn.CisYMDwB iH1_SDate 
            Height          =   345
            Left            =   570
            TabIndex        =   0
            Top             =   360
            Width           =   1815
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
            Object.Width           =   1815
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
         Begin Cis3D_v60.CIS3D PB_SDate 
            Height          =   345
            Left            =   90
            Top             =   360
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   609
            Caption         =   "äJén"
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
            cBoderWidth     =   2
            cFont3D         =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_EDate 
            Height          =   345
            Left            =   90
            Top             =   720
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   609
            Caption         =   "èIóπ"
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
            cBoderWidth     =   2
            cFont3D         =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin CisText_V60.CisText iH1_SBin 
            Height          =   360
            Left            =   2400
            TabIndex        =   1
            Top             =   360
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            cFormat         =   "00"
            cWildeCard      =   -1  'True
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
         Begin CisText_V60.CisText iH1_EBin 
            Height          =   360
            Left            =   2400
            TabIndex        =   3
            Top             =   720
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            cFormat         =   "00"
            cWildeCard      =   -1  'True
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_EDate 
            Height          =   345
            Left            =   570
            TabIndex        =   2
            Top             =   720
            Width           =   1815
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
            Object.Width           =   1815
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
      Begin Cis3D_v60.CIS3D BLB_HINBN 
         Height          =   555
         Left            =   2925
         Top             =   615
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   979
         ForeColor       =   16711680
         Caption         =   "çﬁéø"
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
         cPositionX      =   50
         Begin CisText_V60.CisText iH1_Zai 
            Height          =   360
            Left            =   540
            TabIndex        =   29
            Top             =   120
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
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   555
         Left            =   6675
         Top             =   615
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   979
         ForeColor       =   16711680
         Caption         =   "îwî‘çÜ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         cPositionX      =   50
         Begin CisText_V60.CisText iH1_Seban 
            Height          =   375
            Left            =   720
            TabIndex        =   30
            Top             =   105
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   661
            cWildeCard      =   -1  'True
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
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   375
         Left            =   8580
         Top             =   45
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
         cAlingnment     =   1
         cPositionX      =   30
         Begin CisText_V60.CisText iH1_Ita 
            Height          =   360
            Left            =   555
            TabIndex        =   31
            Top             =   15
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
            cILength        =   3
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   375
         Left            =   8580
         Top             =   420
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
         cAlingnment     =   1
         cPositionX      =   40
         Begin CisText_V60.CisText iH1_Haba 
            Height          =   360
            Left            =   555
            TabIndex        =   32
            Top             =   15
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
            cILength        =   3
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   375
         Left            =   8580
         Top             =   795
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
         cAlingnment     =   1
         cPositionX      =   40
         Begin CisText_V60.CisText iH1_Nagasa 
            Height          =   360
            Left            =   555
            TabIndex        =   33
            Top             =   0
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
            cILength        =   3
         End
      End
   End
   Begin Cis3D_v60.CIS3D UC_3D31 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   714
      BackColor       =   8438015
      ForeColor       =   0
      Caption         =   "Å@Åy çﬁóø•éxããé¿ê—èÛãµè∆âÔ Åz"
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
      Begin VB.PictureBox Back_Over 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   4665
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   45
         Width           =   1995
         Begin VB.OptionButton Op_Overlap 
            BackColor       =   &H00C0C0C0&
            Caption         =   "íPàÍ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   1020
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Value           =   -1  'True
            Width           =   780
         End
         Begin VB.OptionButton Op_Overlap 
            BackColor       =   &H00C0C0C0&
            Caption         =   "èdï°"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   780
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D32 
         Height          =   255
         Left            =   7500
         Top             =   45
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   450
         BackColor       =   8438015
         Caption         =   "íäèoåèêî[              ]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   8438015
         cBoderColor2    =   8438015
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cBoderStyle     =   2
         cPositionX      =   50
         Begin Cis3D_v60.CIS3D lb_Kensu 
            Height          =   285
            Left            =   1095
            Top             =   15
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            BackColor       =   8438015
            ForeColor       =   0
            Caption         =   "##,### / ##,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8438015
            cBoderColor2    =   8438015
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   2
         End
      End
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14415
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
         Left            =   11925
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
   End
   Begin VB.PictureBox Back_Dsp 
      BackColor       =   &H00404040&
      Height          =   720
      Left            =   270
      ScaleHeight     =   660
      ScaleWidth      =   5295
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   450
      Width           =   5355
      Begin Cis3D_v60.CIS3D UC_3D5 
         Height          =   645
         Left            =   15
         Top             =   0
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   1138
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin Cis3D_v60.CIS3D UC_3D9 
            Height          =   570
            Index           =   0
            Left            =   60
            Top             =   30
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   1005
            BackColor       =   16761024
            Caption         =   "ï\é¶êÿë÷"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D UC_3D10 
            Height          =   570
            Left            =   735
            Top             =   30
            Width           =   3510
            _ExtentX        =   6191
            _ExtentY        =   1005
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderStyle     =   1
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "áD"
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   4
               Left            =   2730
               TabIndex        =   13
               Top             =   150
               Width           =   615
            End
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "áC"
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   3
               Left            =   2070
               TabIndex        =   12
               Top             =   150
               Width           =   615
            End
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "áB"
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   2
               Left            =   1395
               TabIndex        =   11
               Top             =   150
               Width           =   615
            End
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "áA"
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   1
               Left            =   720
               TabIndex        =   10
               Top             =   150
               Width           =   615
            End
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H008080FF&
               Caption         =   "á@"
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   0
               Left            =   60
               TabIndex        =   9
               Top             =   150
               Value           =   -1  'True
               Width           =   615
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D7 
            Height          =   570
            Left            =   4260
            Top             =   30
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   1005
            ForeColor       =   16711935
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
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
            Begin VB.OptionButton Op_Sort 
               BackColor       =   &H00C0C0C0&
               Caption         =   "è∏èá"
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Index           =   0
               Left            =   30
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   60
               Value           =   -1  'True
               Width           =   825
            End
            Begin VB.OptionButton Op_Sort 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ç~èá"
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Index           =   1
               Left            =   30
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   315
               Width           =   825
            End
         End
      End
   End
   Begin VB.TextBox Dummy 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   345
      IMEMode         =   3  'µÃå≈íË
      Left            =   240
      TabIndex        =   18
      Text            =   "D"
      Top             =   210
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Left            =   510
      Top             =   210
   End
   Begin VB.Timer Timer2 
      Left            =   1230
      Top             =   210
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   705
      Left            =   8835
      ScaleHeight     =   645
      ScaleWidth      =   5970
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   450
      Width           =   6030
      Begin CisBtn_60.CisBtn PB_Set 
         Height          =   585
         Left            =   2580
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   15
         BtnCaption      =   "ê›íË"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   0
         BtnShadowColor  =   16777215
         PFCaption       =   "( F11 )"
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
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   585
         Left            =   5100
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
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
      Begin CisBtn_60.CisBtn PB_ENT 
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
            Weight          =   700
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
      Begin CisBtn_60.CisBtn PB_CAN 
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
            Weight          =   700
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
      Begin CisBtn_60.CisBtn PB_Csv 
         Height          =   585
         Left            =   1740
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   14
         BtnCaption      =   "èoóÕ"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   192
         BtnShadowColor  =   16777215
         PFCaption       =   "( F10 )"
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
      Begin CisBtn_60.CisBtn PB_Prt 
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   13
         BtnCaption      =   "àÛç¸"
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
         PFCaption       =   "( F8 )"
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   45
         Top             =   30
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   405
      Left            =   0
      Top             =   765
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "CXHZ0610"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :   çﬁóøÅEéxããé¿ê—èÛãµè∆âÔ
'**       ÉtÉHÅ[ÉÄID    :   CXHZ0610
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2008/07/10  By CIS
'**       ïœ  çX  ì˙    :   2014/02/04  By CIS WIN7ëŒâû
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' Ã´∞—±∏√®Ãﬁ Ã◊∏ﬁ
    Dim ProcHB                  As String * 2   ' ì¸óÕàÊêÿë÷
    Dim MaxCnt                  As Long
    Dim NowCnt                  As Long
    Dim mSL_SelectW             As String       ' SQL SELECT(Where)èåèï∂

    Dim CisDBGrid               As New CisDBGrid ' TDBGridPro60∏◊ΩíËã`
    Dim CisComboTehai           As CisCombo
    Dim CisComboHKB             As CisCombo
    Dim CisComboKoku            As CisCombo     '------------ 2008/06/10 ADD
    
    Dim mGotColor               As Long
    Dim mLostColor              As Long


'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås    ÇkÇèÇÅÇÑÉCÉxÉìÉg    Åt              +
'+----------------------------------------------------------+
Private Sub Form_Load()

    If Rv_Call3 = "" Then
    '   #-------------------#
    '   # ìÒ èd ãN ìÆ ñh é~  #
    '   #-------------------#
        If Not Dupli_Start Then End
    '   #-------------------#
    '   # èâ ä˙ ì‡ óe éÊ ìæ  #
    '   #-------------------#
        If Not IniGet Then End
    Else
        FormAct = False
    End If
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    
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

    mGotColor = OpH1_DspKbn(0).BackColor
    mLostColor = OpH1_DspKbn(1).BackColor
    
'+---------------------+
'+ TrueGrid60 èâä˙ê›íË  +
'+---------------------+
    With CisDBGrid
        Set .DBGridObj = DBGrid1
        .FrmCaption = Me.Caption
        .MaxSort = 7                                ' ⁄≤±≥ƒø∞ƒç≈ëÂêî
        .LayOut = True                              ' ⁄≤±≥ƒïœçXèàóù True:óLÇË,False:ñ≥Çµ
        .LayOutMax = 5                              ' ⁄≤±≥ƒèÓïÒç≈ëÂêî
        .LayOutColMax = 30                          ' ⁄≤±≥ƒï\é¶â¬î\çÄñ⁄ç≈ëÂêî
        .LayOutPath = gIniExe                       ' ⁄≤±≥ƒèÓïÒäiî[êÊÅAÃß≤ŸèÍèä
        .LayOutName = CisFun.WSID & "HZ0610.Lay"    ' ⁄≤±≥ƒèÓïÒäiî[êÊÅAÃß≤Ÿñº
        .PrintHeadLine = Ç†ÇË
        .PrintBodyLine = Ç†ÇË
        .PrintBodyGyo = 10
        .CsvName = Trim(Me.Caption) & ".csv"
        .CsvRegKey = gRegKey
        .CsvValue = "CXHZ0610"
        .GridInitGet = False                        ' ∏ﬁÿØƒﬁì‡óeèâä˙ê›íË ∏ÿØÃﬂŒﬁ∞ƒﬁçÏê¨Ã◊∏ﬁ
        .GridProcGet = False
        .Init
    End With
    Call GridInitSet

'   #-------------------#
'   # âÊñ íÜâõà íuï\é¶   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    PNL_ExeMsg.Move 4600, 400

'    If Rv_Call3 = "" Then
        Me.Show
'    End If
End Sub

'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    ÅsÇ`ÇÉÇîÇâÇñÇÅÇîÇÖ ÉCÉxÉìÉgÅt                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
' Headïîèâä˙âª
    B1Op_Patn(0).Value = True
    Op_Sort(0).Value = True
    Op_Sort(0).BackColor = mGotColor
    
    Op_Overlap(0).Value = True
    Op_Overlap(0).BackColor = mGotColor
    
    Call HeadBodyClear("H")
    
' ADO∫›ƒ€∞ŸÅ@ê⁄ë±èÓïÒÅ@ÉZÉbÉg
    Adodc1.ConnectionString = CisDB.ConnectString
    Adodc2.ConnectionString = CisDB.ConnectString
    
    '-----2014/02/04 EDIT START-----
    gStr = CisDB.ConnectString
    gConnect = "server=" & CisDB.Server & ";database=" & CisDB.DBName & ";User ID=" & CisDB.User & ";password=" & CisDB.PassWord(GetSetting(gRegKey, "Environment", "DrowsSap", ""))
    
    gStr = Replace(gStr, "Use Procedure", "Extended Properties=" & gConnect & ";Use Procedure")
    
    Adodc1.ConnectionString = gStr
    Adodc2.ConnectionString = gStr
    '-----2014/02/04 EDIT END-------

    
    If Rv_Call3 = "K" Then
       iH1_SDate = RV_SYmd
       iH1_Zai = RV_Zaist                           ' 2008.07.10 add
       iH1_Seban = RV_Seban
       iH1_Torcd = RV_Torcd
       iH1_Ukeir = RV_Ukeir
       Call iH1_Torcd_LostFocus

       If Not Rv_CallObj2 Is Nothing Then
           Unload Rv_CallObj2
       End If

       Call ReturnPress
    End If
    
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås QueryUnloadÉCÉxÉìÉg Åt                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Rv_Call3 = "" Then
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
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  'ÅyåüçıÅz
           Case vbKeyF8:        If PB_Prt.Visible Then Call PB_Prt_Click    'ÅyàÛç¸Åz
           Case vbKeyF10:       If PB_Csv.Visible Then Call PB_Csv_Click    'ÅyèoóÕÅz
           Case vbKeyF11:       If PB_Set.Visible Then Call PB_Set_Click    'Åyê›íËÅz
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
    If Rv_Call3 = "" Then
    '+---------------------+
    '+   √ﬁ∞¿Õﬁ∞Ωêÿíf
    '+---------------------+
        Call CisDB.DBDISConnect
    '+---------------------+
    '+   Ã´∞—µÃﬁºﬁ™∏ƒè¡ãé
    '+---------------------+
        Call CisFrm.UnLoadDisp
        End
    Else
        Unload CXHZ0610
        If Not Rv_CallObj Is Nothing Then
           Rv_CallObj.Show
        End If
    End If
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
    
    If ProcHB = "H1" Then
       Call HeadBodyClear("H")
        iH1_SDate.SetFocus
        GoTo PB_CAN_Ed
    End If
    
    If ProcHB = "B1" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call DBGridClr
        iH1_SDate.SetFocus
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B2" Then
        ProcHB = "B1"
        Call DispChange(ProcHB)
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    CSVèoóÕ           Å@     +
'+----------------------------+
Private Sub PB_Csv_Click()
    With CisDBGrid
''Call Test_LogOutput("[Adodc1]" + Chr(13) + Chr(10) + Adodc1.RecordSource, 1)    '----------- Debug --------
        Set .Recordset = Adodc1.Recordset
        .PutCsv
    End With
End Sub
'+----------------------------+
'+    àÛç¸             Å@     +
'+----------------------------+
Private Sub PB_Prt_Click()
    With CisDBGrid
        Set .Recordset = Adodc1.Recordset
        .PutPrt
    End With
End Sub
'+----------------------------+
'+    ê›íË              Å@     +
'+----------------------------+
Private Sub PB_Set_Click()
    With CisDBGrid
        .LayOutPatan
        If .LayOutSetFlg Then
           B1Op_Patn(.LayOutSetNo).Value = True
           If ProcHB = "B1" Then
              Call DBInput("R")
           End If
        End If
    End With
End Sub
'+----------------------------+
'+    åü  çı  ÉL Å[(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_SDate" Then iH1_SDate.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_EDate" Then iH1_EDate.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_SSTorcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_UTorcd(0): Exit Sub
    If PB_Look.Tag = "iH1_Seban" Then Call Look_ZaiTori(0): Exit Sub    ' 2008.07.11 upd : Look_Hinbn Å® Look_ZaiTori   ' 2008.07.14 upd : (1) Å® (0)
    If PB_Look.Tag = "iH1_Zai" Then Call Look_ZaiTori(1): Exit Sub      ' 2008.07.11 upd : Look_Hinbn Å® Look_ZaiTori   ' 2008.07.14 upd : iH1_Hinban Å® iH1_Zai, (0) Å® (1)
    If PB_Look.Tag = "iH1_Ita" Then Call Look_ZaiTori(2): Exit Sub      ' 2008.07.14 add
    If PB_Look.Tag = "iH1_Haba" Then Call Look_ZaiTori(3): Exit Sub     ' 2008.07.14 add
    If PB_Look.Tag = "iH1_Nagasa" Then Call Look_ZaiTori(4): Exit Sub   ' 2008.07.14 add
End Sub
'+-----------------------------+
'+       éÊà¯êÊåüçı             +
'+-----------------------------+
Private Sub Look_SSTorcd()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 8
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Torcd = RV_TorRName
    End If
End Sub
'+-----------------------------+
'+       éÊà¯êÊéÛì¸åüçı         +
'+-----------------------------+
Private Sub Look_UTorcd(KB As Integer)
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_Torcd = iH1_Torcd
    RV_TorKb = 8
    RV_Kojyo = ""
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Torcd = RV_TorRName
        iH1_Ukeir = RV_Ukeir
    End If
End Sub

'+-----------------------------+
'+       çﬁóøéÊà¯êÊåüçı        +
'+-----------------------------+
Private Sub Look_ZaiTori(KB As Integer)
    RV_Left = 0
    RV_Top = 0
        
    RV_Call = ""
    RVI_SK = 9
    RVI_Torcd = Trim(iH1_Torcd)
    RVI_Ukeir = Trim(iH1_Ukeir)
    RVI_Seban = RTrim(iH1_Seban)
    
    CKK0077.Show vbModal
    Unload CKK0077
    Set CKK0077 = Nothing

    If RV_Rtn Then
        Select Case KB
            Case 0: iH1_Seban = RV_Seban
            Case 1: iH1_Zai = RV_Zaist
            Case 2:
                iH1_Ita = ""
                If RV_Size = 0 Then
                   If RV_Itatu <> 0 Then iH1_Ita = CisFun.RSetFld(RV_Itatu, 0, iH1_Ita.cFormat)
                End If
            Case 3
                iH1_Haba = ""
                If RV_Size = 0 Then
                   If RV_Width <> 0 Then iH1_Haba = CisFun.RSetFld(RV_Width, 0, iH1_Haba.cFormat)
                End If
            Case 4
                iH1_Nagasa = ""
                If RV_Size = 0 Then
                   If RV_Long <> 0 Then iH1_Nagasa = CisFun.RSetFld(RV_Long, 0, iH1_Nagasa.cFormat)
                End If
        End Select
    End If
End Sub

'*******************************
'*      É^ÉCÉgÉãÉoÅ[ï“èW        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'+-----------------------------+
'+  iH1_SDate
'+-----------------------------+
Private Sub iH1_SDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_SDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_EDate
'+-----------------------------+
Private Sub iH1_EDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_EDate_LostFocus()
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
    H1lb_Torcd = ""
    If TorNmGet(iH1_Torcd, 8) Then
        H1lb_Torcd = TRM_RName
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

'*---------------------------------------------- 2008.07.14 upd start : iH1_Hinbn Å® iH1_Zai
'+-----------------------------+
'+  iH1_Zai
'+-----------------------------+
Private Sub iH1_Zai_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Zai_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'*---------------------------------------------- 2008.07.14 upd end   : iH1_Hinbn Å® iH1_Zai

'*---------------------------------------------- 2008.07.14 add start
'+-----------------------------+
'+  iH1_Ita
'+-----------------------------+
Private Sub iH1_Ita_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Ita_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_Haba
'+-----------------------------+
Private Sub iH1_Haba_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Haba_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_Nagasa
'+-----------------------------+
Private Sub iH1_Nagasa_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Nagasa_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'*---------------------------------------------- 2008.07.14 add end

'+-----------------------------+
'+  iH1_Seban
'+-----------------------------+
Private Sub iH1_Seban_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Seban_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub

'+----------------------------------+
'+      OpH1_DspKbn
'+----------------------------------+
Private Sub OpH1_DspKbn_GotFocus(INDEX As Integer)
    For Each gObj In OpH1_DspKbn
        If gObj.INDEX = INDEX Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
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
    
    H1Mode = False:     H1Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    B2Mode = False:     B2Color = gPLostSel
    
'   ¿ﬁ–∞çÄñ⁄(âÊñ ÇÃÉ`ÉâÉcÉLñhé~)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel: PB_ENT.BtnCaption = "ì¸óÕ"
        Case "B1":      B1Mode = True:  B1Color = gPGotSel: PB_ENT.BtnCaption = "è⁄ç◊"
        Case "B2":      B2Mode = True:  B2Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B2_Area1.Enabled = B2Mode
    B2_Area1.BackColor = B2Color
    
    PB_Look.Visible = H1Mode
    PB_END.Visible = H1Mode
    B2_Area1.Visible = B2Mode
    If Not B2Mode Then
       If PNL_ExeMsg.Visible Then
          PNL_ExeMsg.Visible = False
       End If
    End If
    PB_ENT.Visible = H1Mode Or B1Mode
    PB_Set.Visible = Not B2Mode
    PB_Prt.Visible = B1Mode
    PB_Csv.Visible = B1Mode
    
    If Rv_Call3 = "K" Then
       PB_END.BtnCaption = "ñﬂÇÈ"
       PB_END.Visible = True
'       PB_CAN.Visible = B2Mode
    End If
    
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
            If .Name Like "Op" & wbc_Enm & "*" Then
               .Value = False
               .BackColor = mLostColor
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "*Cmb*" Then .ListIndex = 0
        End With
    Next gObj
    If HeadBody = "H" Then
        iH1_SDate = Format(Date, "YYYYMMDD")
        OpH1_DspKbn(0).Value = True
        OpH1_DspKbn(0).BackColor = mGotColor
    End If
End Sub
'****************************
'*      ì¸óÕèàóù(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
       If Not Head1Chk Then
          GoTo ReturnPress_Ed
       End If
       ProcHB = "B1"
       Call DispChange(ProcHB)
       DBGrid1.SetFocus
       GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
       Call BodySet
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
    
    ' ì˙ït
    If Trim(iH1_SDate) <> "" Then
       If CisFun.Val2(iH1_SDate.cDay) = 0 Then
          iH1_EDate = ""
          iH1_EBin = ""
       End If
    End If
    If Trim(iH1_SDate) <> "" And Trim(iH1_EDate) <> "" Then
       If iH1_SDate & Format(iH1_SBin, "00") > iH1_EDate & Format(iH1_EBin, "00") Then
            With CisFun
                .MB_Lines = 5
                   .MB_MSG(2) = "    ó\íËì˙ÅEï÷Çê≥ÇµÇ≠ì¸óÕÇµÇƒâ∫Ç≥Ç¢         "
                .MB_MSG(4) = "        äJénÅ@ÅÑ   èIóπ             "
                .MB_Button = Error
                .MBOX
            End With
            iH1_SDate.SetFocus
            Exit Function
       End If
    End If
    ' î[ì¸î‘çÜ
    If InStr(1, iH1_SNouNo, "*") > 0 Then
       iH1_ENouNo = ""
    End If
    
    If Not DBInput Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    äYìñÉfÅ[É^Ç™ë∂ç›ÇµÇ‹ÇπÇÒ°            "
            .MB_MSG(3) = "    çƒìxíäèoèåèÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢       "
            .MB_Title = "íäèoèàóù"
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If
    
    Head1Chk = True
End Function
'+----------------------------+
'+    ï\é¶êÿë÷µÃﬂºÆ›Œﬁ¿›Å@     +
'+----------------------------+
Private Sub B1Op_Patn_Click(INDEX As Integer)
    Call CisFrm.MousePT(11)

    For Each gObj In B1Op_Patn
        If INDEX = gObj.INDEX Then
           gObj.BackColor = &H8080FF
        Else
           gObj.BackColor = &HC0C0C0
        End If
    Next gObj

    CisDBGrid.LayOutNow = INDEX + 1
    CisDBGrid.LayOutDisp

    If ProcHB = "B1" Then
        Call DBInput("R")
    End If
    Call CisFrm.MousePT(1)

End Sub
'+---------------------------+
'+    èdï°ï\é¶Å@¡™Ø∏ŒﬁØ∏Ω     +
'+---------------------------+
Private Sub Op_Overlap_GotFocus(INDEX As Integer)
    For Each gObj In Op_Overlap
        If gObj.INDEX = INDEX Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
Private Sub Op_Overlap_Click(INDEX As Integer)

    If Not FormAct Then Exit Sub
    
    Call CisFrm.MousePT(11)

    Call CisDBGrid.MargeDisp(Op_Overlap(1).Value)

    If ProcHB = "H1" Then
        iH1_SDate.SetFocus
    Else
        DBGrid1.SetFocus
    End If

    Call CisFrm.MousePT(1)
End Sub
'+---------------------------+
'+    ï™óﬁï\é¶Å@¡™Ø∏ŒﬁØ∏Ω     +
'+---------------------------+
Private Sub Op_Sort_GotFocus(INDEX As Integer)
    For Each gObj In Op_Sort
        If gObj.INDEX = INDEX Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
Private Sub Op_Sort_Click(INDEX As Integer)
    If ProcHB = "H1" Then Exit Sub
    Call DBInput
End Sub
'**************************************
'*                                    *
'*       ÉfÅ[É^ÉxÅ[ÉXì¸óÕ Åï ï\é¶     *
'*                                    *
'**************************************
Private Function DBInput(Optional InType As String = "") As Boolean
    Dim DateStr As String
    Dim DateSt  As String
    Dim DateEd  As String
    Dim BinSt   As String
    Dim BinEd   As String
    
    DBInput = True
    
    Call CisFrm.MousePT(11)
    
    On Error GoTo DBInput_Err:
' íäèoíÜÉÅÉbÉZÅ[ÉWï\é¶
    PNL_Wait.Visible = True
    PNL_PF.Enabled = False
    DoEvents
   
' èâä˙ï\é¶éûÅAèåèï∂çÏê¨(Where)
    If ProcHB = "H1" And InType = "" Then
'       ï\é¶èåè ÉZÉbÉg
       CisDBGrid.Select_Where = ""
       With CisFun
            .Where_CreateKB = NewDocument
            ' ì˙ït
            DateStr = "é¿ê—"                        ' 2008.07.11 add
            DateSt = "": DateEd = "": BinSt = "": BinEd = ""
            ' ì˙ït åééwíË
            If Trim(iH1_SDate) <> "" And CisFun.Val2(iH1_SDate.cDay) = 0 Then
               .Where_Create AlphaMe, "Substring(" & DateStr & "ì˙,1,6)", iH1_SDate
               .Where_Create AlphaMe, DateStr & "ï÷", Format(iH1_SBin, "00")
            End If
            ' ì˙ïtäJénï“èW
            If Trim(iH1_SDate) <> "" Then
               DateSt = iH1_SDate
            End If
            If CisFun.Val2(iH1_SBin) <> 0 Then
               BinSt = Format(iH1_SBin, "00")
            End If
            ' ì˙ïtèIóπï“èW
            If Trim(iH1_EDate) <> "" Then
               DateEd = iH1_EDate
            End If
            If CisFun.Val2(iH1_EBin) <> 0 Then
               BinEd = Format(iH1_EBin, "00")
            End If
            ' ì˙ïtäJénÇÃÇ›
            If Trim(DateSt) <> "" And Trim(DateEd) = "" And CisFun.Val2(iH1_SDate.cDay) <> 0 Then
               If Trim(BinSt) = "" Then
                  .Where_Create AlphaMe, DateStr & "ì˙", DateSt
               Else
                  .Where_Create AlphaMe, DateStr & "ì˙+" & DateStr & "ï÷", DateSt & BinSt
               End If
            End If
            ' ì˙ïtèIóπÇÃÇ›
            If Trim(DateSt) = "" And Trim(DateEd) <> "" Then
               If Trim(BinEd) = "" Then
                  .Where_Create AlphaMe, DateStr & "ì˙", DateEd, à»â∫
               Else
                  .Where_Create AlphaMe, DateStr & "ì˙+" & DateStr & "ï÷", DateEd & BinEd, à»â∫
               End If
            End If
            ' ì˙ïtäJénÅEèIóπ
            If Trim(DateSt) <> "" And Trim(DateEd) <> "" Then
               If Trim(BinSt) = "" And Trim(BinEd) = "" Then
                  .Where_Create AlphaMe, DateStr & "ì˙", DateSt, Between, , DateEd
               Else
                  If Trim(BinSt) = "" Then BinSt = "00"
                  If Trim(BinEd) = "" Then BinEd = "99"
                  .Where_Create AlphaMe, DateStr & "ì˙+" & DateStr & "ï÷", DateSt & BinSt, Between, , DateEd & BinEd
               End If
            End If
            ' éxããêÊÅié¿ê—êÊÅj
            .Where_Create AlphaMe, "éxããêÊ", iH1_Torcd
            ' éÛì¸
            .Where_Create AlphaMe, "éÛì¸", iH1_Ukeir
            ' çﬁéø                                          ' 2008.07.11 add
            .Where_Create AlphaMe, "çﬁéø", iH1_Zai          ' 2008.07.11 add
            ' îwî‘çÜ
            .Where_Create AlphaMe, "îwî‘çÜ", iH1_Seban
            '----------------------------------------------------- 2008.07.14 add start
            ' î¬å˙
            If .Val2(iH1_Ita) <> 0 Then
               .Where_Create Numeric, "î¬å˙", .Val2(iH1_Ita)
            End If
            ' ïù
            If .Val2(iH1_Haba) <> 0 Then
               .Where_Create Numeric, "ïù", .Val2(iH1_Haba)
            End If
            ' í∑Ç≥
            If Trim(iH1_Nagasa) <> "" Then
               .Where_Create Numeric, "í∑Ç≥", .Val2(iH1_Nagasa), ìôÇµÇ¢, Create
            End If
            '----------------------------------------------------- 2008.07.14 add end
            ' î[î‘
            If Trim(iH1_SNouNo) <> "" And Trim(iH1_ENouNo) <> "" Then
               .Where_Create AlphaMe, "î[ì¸î‘çÜ", iH1_SNouNo, Between, , iH1_ENouNo
            Else
               If Trim(iH1_SNouNo) <> "" Then
                  .Where_Create AlphaMe, "î[ì¸î‘çÜ", iH1_SNouNo
               End If
               If Trim(iH1_ENouNo) <> "" Then
                  .Where_Create AlphaMe, "î[ì¸î‘çÜ", iH1_ENouNo, à»â∫
               End If
            End If
'*--------------------------------------------------------------------------' 2008.07.11 add start
            ' ÉfÅ[É^ãÊï™
            ' éxãã (çﬁóøé¿ê—ä«óùÉeÅ[ÉuÉã[ì¸óÕéÌóﬁ]=ãÛîí(NULLä‹Çﬁ))
            If OpH1_DspKbn(1).Value Then
                .Where_Create AlphaMe, "ì¸óÕéÌóﬁÉRÅ[Éh", 1, ìôÇµÇ¢, Create
            End If
            ' íºëóéxãã (çﬁóøé¿ê—ä«óùÉeÅ[ÉuÉã[ì¸óÕéÌóﬁ]=A)
            If OpH1_DspKbn(2).Value Then
                .Where_Create AlphaMe, "ì¸óÕéÌóﬁÉRÅ[Éh", 2, ìôÇµÇ¢, Create
            End If
            ' é¿ê—ãÊï™ (3)
            .Where_Create AlphaMe, "é¿ê—ãÊï™", 3, ìôÇµÇ¢, Create
'*--------------------------------------------------------------------------' 2008.07.11 add end
            
            CisDBGrid.Select_Where = .Where_Phrase
Call Test_LogOutput("ì¸óÕ[åüçıèåè01]" + Chr(13) + Chr(10) + .Where_Phrase, 0)  '<<<<<------------ for debug
       End With
    End If
    
' ⁄≤±≥ƒï\é¶ì‡óeÇÊÇË√ﬁ∞¿òAåãñºÇãÅÇﬂÇÈ
    CisDBGrid.LayOutGet
    
    gSL_Select = "select " & CisDBGrid.LayOutfields
    gSL_Select = gSL_Select & ", é¿ê—ì˙, é¿ê—ï÷, î[ïièëï≈, î[ïièëçs, î[ì¸î‘çÜ, îwî‘çÜ,"
    gSL_Select = gSL_Select & " éxããêÊ, ó™èÃ, çﬁóøä«óùî‘çÜ"
    gSL_Select = gSL_Select & " from çﬁóøéxããé¿ê—èÛãµñ‚çáÉrÉÖÅ[ "
Call Test_LogOutput("ì¸óÕ[åüçıèåè02]" + Chr(13) + Chr(10) + gSL_Select, 0)     '<<<<<------------ for debug

' ï™óﬁì‡óe(Order ByãÂ) æØƒ
    With CisDBGrid
        .OrderSet (Op_Sort(0).Value)
        If .Select_OrderBy = "" Then
            DBInput = False
            GoTo DBInput_Ed
        End If
    End With
' íäèoï∂Å@çÏê¨
    gSL_Select = gSL_Select & CisDBGrid.Select_Where & CisDBGrid.Select_OrderBy
Call Test_LogOutput("ì¸óÕ[åüçıèåè03]" + Chr(13) + Chr(10) + gSL_Select, 0)     '<<<<<------------ for debug

' ÉOÉäÉbÉhì‡ï\é¶
    If Not DBGridDisp Then
        DBInput = False
        GoTo DBInput_Ed
    End If
    On Error Resume Next
' ï\é¶ÉfÅ[É^Ç»Çµ
    On Error GoTo 0
    If Adodc1.Recordset.EOF Then
        DBInput = False
        GoTo DBInput_Ed
    End If
    
' íäèoåèêîï\é¶
    lb_Kensu.Visible = True
    Call KensuDisp
    
' ÉJÉâÉÄÅFçﬁóøñºÇîÒï\é¶Ç…Ç∑ÇÈ                  ' 2008.07.14 add
    DBGrid1.Columns.Item(21).Visible = False    ' 2008.07.14 add
    
    On Error GoTo 0
DBInput_Ed:
    Call CisFrm.MousePT(0)
    
    PNL_Wait.Visible = False
    PNL_PF.Enabled = True
    SendKeys "{END}"
    SendKeys "{RIGHT}"
    SendKeys "{LEFT}"
    SendKeys "{HOME}"
    Exit Function
DBInput_Err:
    Call CisFrm.MousePT(0)
    CisFun.ErrorBox
    PNL_Wait.Visible = False
    PNL_PF.Enabled = True
End Function
'+----------------------------------------------------------+
'+                                                          +
'+                                                          +
'+      ÉOÉäÉbÉhèàóù(TrueDBGrid)                             +
'+                                                          +
'+                                                          +
'+----------------------------------------------------------+
'+----------------------------------------------------------+
'+   ÉOÉäÉbÉhì‡óeèâä˙ê›íË
'+----------------------------------------------------------+
Private Sub GridInitSet()
' Call .GridInitSet(œ∞ºﬁèàóù,ï™óﬁèàóù,√ﬁ∞¿òAåãñº,àÛç¸å©èoÇµ,àÛç¸åÖêî,çáåvàÛç¸)
    With CisDBGrid
         Call .GridInitSet(True, , "é¿ê—ì˙ï÷", "é¿ê—ì˙ - ï÷", 15)
         Call .GridInitSet(True, , "ï\é¶î[ì¸î‘çÜ", "î[  î‘", 9)
         Call .GridInitSet(True, , "çﬁéø", "çﬁ     éø", 30)
         Call .GridInitSet(True, , "ï\é¶îwî‘çÜ", "îwî‘çÜ", 8)
         Call .GridInitSet(, , "é˚óeêî", "é˚óeêî", 6, False)
         Call .GridInitSet(, , "é¿ê—ñáêî", "é¿ê—ñáêî", 8, False)
         Call .GridInitSet(, , "é¿ê—êî", "é¿ê—êî", 10, False)
         Call .GridInitSet(, , "íPà ñº", "íPà ", 13)
         Call .GridInitSet(True, , "ï\é¶ê°ñ@", "ê°   ñ@", 24)
         Call .GridInitSet(True, , "ï“èWéxããêÊ", "éxããêÊ", 25)
         Call .GridInitSet(True, , "éÛì¸", "éÛ", 3)
         Call .GridInitSet(True, , "çﬁéÌÉRÅ[Éh", "çﬁéÌÉRÅ[Éh", 15)
         Call .GridInitSet(True, , "ÉfÅ[É^ãÊï™ñº", "ÉfÅ[É^ãÊï™", 8)
         Call .GridInitSet(True, , "çﬁéÌñº", "çﬁéÌ", 12)
         Call .GridInitSet(True, , "å`èÛñº", "å`èÛ", 12)
         Call .GridInitSet(True, , "íuèÍ", "íuèÍ", 11)
         Call .GridInitSet(True, , "ÉTÉCÉNÉã", "ÉTÉCÉNÉã", 8)
         Call .GridInitSet(True, , "ï\é¶óLñ≥èûãÊï™", "óLñ≥", 4)
         Call .GridInitSet(, , "é¿ê—ä∑éZíl", "é¿ê—ä∑éZíl", 13)          'åÖêîOKÅH
         Call .GridInitSet(, , "íPâø", "íP  âø", 13)
         Call .GridInitSet(, , "ã‡äz", "ã‡  äz", 13)
    End With
End Sub
'+----------------------------------------------------------+
'+      ADODC1(åèêîï\é¶)                                     +
'+----------------------------------------------------------+
Private Sub KensuDisp()
    Static NowCnt       As Long
    Static MaxCnt       As Long
    NowCnt = Adodc1.Recordset.AbsolutePosition
    MaxCnt = Adodc1.Recordset.RecordCount
    If NowCnt < 0 Then Exit Sub
    lb_Kensu.Caption = Format(NowCnt, "###,##0") & " / " & _
                       Format(MaxCnt, "###,##0")

End Sub
'+---------------------+
'+  ÉOÉäÉbÉhì‡óeï\é¶    +
'+---------------------+
Private Function DBGridDisp()
    
    DBGridDisp = False
    
    On Error GoTo DBGridDisp_ER:
    
    On Error Resume Next

    Adodc1.RecordSource = gSL_Select
    DBGrid1.ReBind
    Adodc1.Refresh
    DBGrid1.Refresh
    
' èdï°ï\é¶
    Call CisDBGrid.MargeDisp(Op_Overlap(1).Value)
    Call KensuDisp

    With DBGrid1
         .SetFocus
    End With
    SendKeys "{HOME}"
    
    DBGridDisp = True
    
    On Error GoTo 0
    Exit Function
DBGridDisp_ER:
    If Err <> 40002 Then        ' RDO É^ÉCÉÄÉAÉEÉg
       CisFun.ErrorBox
    End If
End Function
'+-----------------------+
'+  ÉOÉäÉbÉhì‡óeèâä˙âª    +
'+-----------------------+
Private Sub DBGridClr()

    Call CisDBGrid.SplitsClr("", 0)   ' ï™äÑÕﬂ≤›âèú
    
    Adodc1.Recordset.Close
    DBGrid1.Close True
    
End Sub
'+----------------------------------------------------------+
'+ ÉLÅ[ì¸óÕÇ…Çƒçsà⁄ìÆ Ås KeyUp ÉCÉxÉìÉg   Åt                  +
'+----------------------------------------------------------+
Private Sub DBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or _
        KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then
        Call KensuDisp
        If B2_Area1.Visible Then
            Call BodySet
        End If
    Else
        Exit Sub
    End If
End Sub
'+----------------------------------------------------------+
'+ ï\é¶çÄñ⁄à⁄ìÆå„ ï™óﬁï\é¶ Ås ColMove ÉCÉxÉìÉg   Åt            +
'+----------------------------------------------------------+
Private Sub DBGrid1_ColMove(ByVal Position As Integer, Cancel As Integer)
    CisDBGrid.ColMoveFlg = True
End Sub
'+----------------------------------------------------------+
'+ ï\é¶çÄñ⁄à⁄ìÆå„ ï™óﬁï\é¶ Ås Paint ÉCÉxÉìÉg   Åt              +
'+----------------------------------------------------------+
Private Sub DBGrid1_Paint()
    If Not FormAct Then Exit Sub
    If CisDBGrid.ColMoveFlg Then  ' çƒï\é¶(ï™óﬁ)
        CisDBGrid.ColMoveFlg = False
        Call DBInput
    End If
End Sub
'+----------------------------------------------------------+
'+ ñæç◊âÊñ ï\é¶(B2_AREA1)
'+----------------------------------------------------------+
Private Sub DBGrid1_DblClick()
    If Not B2_Area1.Visible Then
        Call BodySet
    End If
End Sub
'+----------------------------------------------------------+
'+      ÉOÉäÉbÉh    Ås MouseDown  ÉCÉxÉìÉg   Åt               +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DBGrid1.RowContaining(Y) = -1 And _
        DBGrid1.ColContaining(X) = -1 Then
        Call CisDBGrid.SplitsClr("C", 0)          ' ï™äÑÕﬂ≤›ÇÃÇ›âèú
        Call DBInput
        Exit Sub
    End If
    If DBGrid1.RowContaining(Y) < 0 Then        ' ï™äÑÕﬂ≤›ï\é¶
        If Button = 2 Then
            Call CisDBGrid.SplitsAdd(DBGrid1.ColContaining(X))
            Call DBInput
        End If
        Exit Sub
    End If
    If DBGrid1.RowContaining(Y) < 0 Then Exit Sub
    If Button = 2 Then
        DBGrid1.Row = DBGrid1.RowContaining(Y)
        If B2_Area1.Visible Then
            Call BodySet
        End If
'        If DBGrid1.RowContaining(X) < 0 Then Exit Sub
        Call TotalDsp(DBGrid1.ColContaining(X))
    Else
        DBGrid1.Row = DBGrid1.RowContaining(Y)
        If B2_Area1.Visible Then
            Call BodySet
        End If
    End If
    
    If Button = 2 Then
        DBGrid1.Row = DBGrid1.RowContaining(Y)
        If B2_Area1.Visible Then
            Call BodySet
        End If
'        If DBGrid1.RowContaining(X) < 0 Then Exit Sub
        Call TotalDsp(DBGrid1.ColContaining(X))
    Else
        If PNL_ExeMsg.Visible Then
           PNL_ExeMsg.Visible = False
        End If
        DBGrid1.Row = DBGrid1.RowContaining(Y)
        If B2_Area1.Visible Then
            Call BodySet
        End If
    End If
End Sub
'+----------------------------------------------------------+
'+      ÉOÉäÉbÉh    Ås MouseUp  ÉCÉxÉìÉg   Åt                +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call KensuDisp
    If Button = 2 Then
       PNL_ExeMsgX = ""
       lb_THSu = "": lb_TNSu = "": lb_TNYSu = ""
       PNL_ExeMsg.Visible = False
    End If
End Sub
'**********************************************************
'*  ñæç◊ï\é¶
'**********************************************************
Private Sub BodySet()
    Dim ICnt   As Integer
    ProcHB = "B2"
    B2_Area1.Visible = True
    PB_Csv.Visible = False
    PB_Prt.Visible = False
    PB_Set.Visible = False
    PB_ENT.Visible = False
    PB_CAN.Visible = True
        
    For ICnt = 0 To 21
        B2lb_Mei(ICnt) = RTrim(DBGrid1.Columns(ICnt))
    Next ICnt
    For Each gObj In B2lb_Mei
        gObj.Caption = ""
    Next gObj
    gSL_Select = "Select * From çﬁóøéxããé¿ê—èÛãµñ‚çáÉrÉÖÅ[ "
    gSL_Select = gSL_Select & " Where çﬁóøä«óùî‘çÜ = " & Adodc1.Recordset("çﬁóøä«óùî‘çÜ")
    If ZSVRead(gSL_Select, 1) Then
       With ZSV
            B2lb_Mei(0) = .é¿ê—ì˙ï÷
            B2lb_Mei(1) = .ï\é¶î[ì¸î‘çÜ
            B2lb_Mei(2) = .çﬁéø
            B2lb_Mei(3) = Trim(.ï\é¶îwî‘çÜ)
            B2lb_Mei(4) = Format(.é˚óeêî, DBGrid1.Columns(4).NumberFormat)
            B2lb_Mei(5) = Format(.é¿ê—ñáêî, DBGrid1.Columns(5).NumberFormat)
            B2lb_Mei(6) = Format(.é¿ê—êî, DBGrid1.Columns(6).NumberFormat)
            B2lb_Mei(7) = .íPà ñº
            B2lb_Mei(8) = .ï\é¶ê°ñ@
            B2lb_Mei(9) = .ï“èWéxããêÊ
            B2lb_Mei(10) = .éÛì¸
            B2lb_Mei(11) = .çﬁéÌÉRÅ[Éh
            B2lb_Mei(12) = .ÉfÅ[É^ãÊï™ñº
            B2lb_Mei(13) = .çﬁéÌñº
            B2lb_Mei(14) = .å`èÛñº
            B2lb_Mei(15) = .íuèÍ
            B2lb_Mei(16) = .ÉTÉCÉNÉã
            B2lb_Mei(17) = Mid(.ï\é¶óLñ≥èûãÊï™, 1, 2)
            B2lb_Mei(18) = Format(.é¿ê—ä∑éZíl, "###,##0.000")
            B2lb_Mei(19) = Format(.íPâø, "###,##0.000")
            B2lb_Mei(20) = Format(.ã‡äz, "#,###,##0.000")
            B2lb_Mei(21) = .çﬁóøñº
       End With
    End If

GoTo jmp1   ' Ç±ÇÃâ∫ÇÃèàóùÇÕïsóvÇ∆évÇÌÇÍÇÈÇ™ämíËÇ∑ÇÈÇ‹Ç≈écÇµÇƒâÒîÇ∑ÇÈÅ@2008.07.11 ---------------------------------------
    
    On Error Resume Next
    Adodc2.Recordset.Close
'    TDBGrid1.Close True    ' 2008.07.11 del
    
    gSL_Select = "Select * From î≠íçì¸â◊èÛãµñ‚çáÉrÉÖÅ["
    gSL_Select = gSL_Select & " Where ïiî‘ = '" & Adodc1.Recordset("ïiî‘") & "'"
    gSL_Select = gSL_Select & "   And édì¸êÊ = '" & Adodc1.Recordset("édì¸êÊ") & "'"
    gSL_Select = gSL_Select & "   And éÛì¸ = '" & Adodc1.Recordset("éÛì¸") & "'"
    gSL_Select = gSL_Select & "   And Substring(î[ì¸î‘çÜ,1,4) = '" & Mid(Adodc1.Recordset("î[ì¸î‘çÜ"), 1, 4) & "'"
    gSL_Select = gSL_Select & " Order By î[ì¸î‘çÜ "
    Adodc2.RecordSource = gSL_Select
'    TDBGrid1.ReBind        ' 2008.07.11 del
    Adodc2.Refresh
'    TDBGrid1.Refresh       ' 2008.07.11 del
    If Adodc2.Recordset.RecordCount > 1 Then
'       B2_Area2.Visible = True                 ' 2008.07.10 del
       Do Until Adodc2.Recordset.EOF
            
          If Adodc1.Recordset("î[ì¸î‘çÜ") = Adodc2.Recordset("î[ì¸î‘çÜ") Then
             Exit Do
          End If
          Adodc2.Recordset.MoveNext
       Loop
    Else
'       B2_Area2.Visible = False                ' 2008.07.10 del
    End If
jmp1:
End Sub

'+----------------------+
'+      èWåvï\é¶Å@       +
'+----------------------+
Private Sub TotalDsp(JCol As Long)
    Dim wWhere          As String
    Dim wMsg            As String
    Dim wValue          As String
    
    gSL_Select = "Select Sum(é¿ê—êî), Sum(é¿ê—ñáêî), Sum(ã‡äz), "
    gSL_Select = gSL_Select & " Count(*) "
    gSL_Select = gSL_Select & " From çﬁóøéxããé¿ê—èÛãµñ‚çáÉrÉÖÅ[ "
'Call Test_LogOutput("èWåv" + Chr(13) + Chr(10) + gSL_Select, 1)  '<<<<<--------------------------------------<<<<<< for debug
    
    Select Case JCol
      Case -1
           wWhere = ""
      Case 0, 2, 3, 9, 11, 12, 13, 14, 15, 16
                    ' î[ì¸ì˙ÅEï÷(0), çﬁéø(2), îwî‘çÜ(3), éxããêÊ(9), çﬁéÌÉRÅ[Éh(11), ÉfÅ[É^ãÊï™(12), çﬁéÌ(13),
                    ' å`èÛ(14),íuèÍ(15),ÉTÉCÉNÉã(16)
           wMsg = DBGrid1.Columns(JCol).Caption
           wValue = DBGrid1.Columns(JCol).Text
           wWhere = DBGrid1.Columns(JCol).DataField & " = '" & RTrim(DBGrid1.Columns(JCol).Value) & "'"
      Case 1        ' î[ì¸î‘çÜ
           wMsg = DBGrid1.Columns(JCol).Caption
           wValue = DBGrid1.Columns(JCol).Text
           wWhere = "Substring(î[ì¸î‘çÜ, 1, 4) = '" & Mid(DBGrid1.Columns(JCol).Text, 1, 4) & "'"
      Case 10       ' édì¸êÊ•éÛì¸
           wMsg = "édì¸êÊ•éÛì¸"
           wValue = DBGrid1.Columns(JCol - 1).Text & "-[" & DBGrid1.Columns(JCol).Text & "]"
           If wValue = "-[]" Then wValue = ""
           wWhere = DBGrid1.Columns(JCol - 1).DataField & " + " & DBGrid1.Columns(JCol).DataField & " = '" & _
                    DBGrid1.Columns(JCol - 1).Text & DBGrid1.Columns(JCol).Text & "'"
       Case Else
            Exit Sub
    End Select

    wMsg = Replace(wMsg, "Å@", "")
    wMsg = Replace(wMsg, " ", "")
    If JCol < 0 Then
        PNL_ExeMsgX = "Åy èWåvíPà  ÅÅ ëSåèëŒè€ Åz"
        lb_TMei.Caption = ""
    Else
        PNL_ExeMsgX = "Åy èWåvíPà  ÅÅ " & wMsg & " Åz"
        lb_TMei.Caption = wMsg & " ÅF " & wValue
    End If

    If wWhere <> "" Then
        If Trim(CisDBGrid.Select_Where) = "" Then
           gSL_Select = gSL_Select & " Where " & wWhere
        Else
           gSL_Select = gSL_Select & CisDBGrid.Select_Where
           gSL_Select = gSL_Select & " And " & wWhere
        End If
    Else
        gSL_Select = gSL_Select & CisDBGrid.Select_Where
    End If

    With CisDB
        .SQL = gSL_Select
        If .DBRead(1, 0) Then
            lb_THSu = Format(.Recordset(0), "###,##0.000")
            lb_TNSu = Format(.Recordset(1), "###,##0")
            lb_TNYSu = Format(.Recordset(2), "###,##0.000")
            lb_SKensu = " Åy èWåvåèêî ÅÅ " & Format(.Recordset(3), "###,##0") & " Åz"
        Else
            lb_SKensu = "Åy èWåvåèêî ÅÅ   0 Åz"
        End If
        On Error Resume Next
            .ConnectNo = 0
            .Recordset.Close
        On Error GoTo 0
    End With

    PNL_ExeMsg.Visible = True
'
End Sub

''+-------------------------------------+
''+      Debugóp SQLãLò^                +
''+-------------------------------------+
Private Sub Test_LogOutput(str As String, Msg As Integer)
Exit Sub   '=========================================== é¿çsÇµÇ»Ç¢Ç∆Ç´ÇÕÇ±ÇÃçsÇê∂Ç©Ç∑Ç±Ç∆ =========================
    Dim dt01 As String
    Dim dt02 As String
    Dim f_No As Integer
    
    If Msg <> 0 Then MsgBox (str)
    dt01 = CStr(DateTime.Now)
    dt02 = Mid(dt01, 1, 4) & "." & _
           Mid(dt01, 6, 2) & "." & _
           Mid(dt01, 9, 2)

    f_No = FreeFile                             'ÉtÉ@ÉCÉãî‘çÜÇÃéÊìæ
    Open "D:\TestLog\SQL_" & dt02 & ".log" _
          For Append As #f_No                   'ÉtÉ@ÉCÉãÇí«â¡ÉÇÅ[ÉhÇ≈äJÇ≠
    Print #f_No, "[" & Me.Name & "] " & dt01 & _
                 " ----------------------------------------" & Chr(13) & Chr(10) & _
                 str & Chr(13) & Chr(10)        'ÉtÉ@ÉCÉãÇ÷èëÇ´çûÇﬁ"
    Close #f_No                                 'ÉtÉ@ÉCÉãÇï¬Ç∂ÇÈ
End Sub
