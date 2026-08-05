VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CXM0040 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "ÉJÉåÉìÉ_É}ÉXÉ^  ÉÅÉìÉeÉiÉìÉX"
   ClientHeight    =   11130
   ClientLeft      =   1485
   ClientTop       =   1440
   ClientWidth     =   15240
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ÇlÇr ÉSÉVÉbÉN"
      Size            =   12
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z µ∞¿ﬁ∞
   ScaleHeight     =   11130
   ScaleWidth      =   15240
   WindowState     =   2  'ç≈ëÂâª
   Begin Cis3D_v60.CIS3D PB_Auto 
      Height          =   615
      Left            =   13440
      Top             =   1755
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1085
      ForeColor       =   12582912
      Caption         =   "é©ìÆçÏê¨"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr Çoñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      cAlingnment     =   7
      cButton         =   -1  'True
   End
   Begin VB.PictureBox B2_Area1 
      Height          =   8340
      Left            =   6585
      ScaleHeight     =   8280
      ScaleWidth      =   8535
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2550
      Width           =   8595
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
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
         Height          =   8070
         Index           =   0
         Left            =   60
         ScaleHeight     =   8070
         ScaleWidth      =   8415
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   90
         Width           =   8415
         Begin Cis3D_v60.CIS3D P_TenKey 
            Height          =   1110
            Left            =   5265
            Top             =   6945
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   1958
            ForeColor       =   128
            Caption         =   "<ÉeÉìÉLÅ[>"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr Çoñæí©"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cPositionX      =   30
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   285
               Index           =   0
               Left            =   240
               Top             =   480
               Width           =   2730
               _ExtentX        =   4815
               _ExtentY        =   503
               ForeColor       =   0
               Caption         =   "ç∂éÛì¸:[4]Å@ÅEÅ@[6]:âEéÛì¸"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   12632256
               cBoderColor2    =   12632256
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cBoderStyle     =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   285
               Index           =   1
               Left            =   1410
               Top             =   810
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   503
               ForeColor       =   0
               Caption         =   "[2]:éüåé"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   12632256
               cBoderColor2    =   12632256
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cBoderStyle     =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   285
               Index           =   2
               Left            =   1395
               Top             =   150
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   503
               ForeColor       =   0
               Caption         =   "[8]:ëOåé"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   12632256
               cBoderColor2    =   12632256
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cBoderStyle     =   2
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   600
            Left            =   4815
            Top             =   0
            Width           =   3585
            _ExtentX        =   6324
            _ExtentY        =   1058
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
            Begin Cis3D_v60.CIS3D Nisu_Msg 
               Height          =   510
               Index           =   3
               Left            =   2280
               Top             =   45
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   900
               Caption         =   "â“ìÆ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   30
               Begin Cis3D_v60.CIS3D B1lb_Nisu 
                  Height          =   510
                  Index           =   2
                  Left            =   540
                  Top             =   0
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   900
                  BackColor       =   14737632
                  Caption         =   "99.9"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  cFont3DColor2   =   8421504
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin Cis3D_v60.CIS3D Nisu_Msg 
               Height          =   510
               Index           =   0
               Left            =   30
               Top             =   45
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   900
               Caption         =   "íããŒ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   75
               Begin Cis3D_v60.CIS3D B1lb_Nisu 
                  Height          =   510
                  Index           =   0
                  Left            =   675
                  Top             =   0
                  Width           =   465
                  _ExtentX        =   820
                  _ExtentY        =   900
                  BackColor       =   14737632
                  Caption         =   "99"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  cFont3DColor2   =   8421504
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
            Begin Cis3D_v60.CIS3D Nisu_Msg 
               Height          =   510
               Index           =   1
               Left            =   1200
               Top             =   45
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   900
               Caption         =   "ñÈãŒ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   75
               Begin Cis3D_v60.CIS3D B1lb_Nisu 
                  Height          =   510
                  Index           =   1
                  Left            =   600
                  Top             =   0
                  Width           =   465
                  _ExtentX        =   820
                  _ExtentY        =   900
                  BackColor       =   14737632
                  Caption         =   "99"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  cFont3DColor2   =   8421504
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
            End
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   1
            Left            =   1200
            Top             =   1770
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D LB_Week 
            Height          =   600
            Index           =   1
            Left            =   1200
            Top             =   630
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   1058
            BackColor       =   8421504
            Caption         =   "åé"
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
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D LB_Week 
            Height          =   600
            Index           =   2
            Left            =   2400
            Top             =   630
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   1058
            BackColor       =   8421504
            Caption         =   "âŒ"
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
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D LB_Week 
            Height          =   600
            Index           =   3
            Left            =   3600
            Top             =   630
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   1058
            BackColor       =   8421504
            Caption         =   "êÖ"
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
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D LB_Week 
            Height          =   600
            Index           =   4
            Left            =   4800
            Top             =   630
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   1058
            BackColor       =   8421504
            Caption         =   "ñÿ"
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
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D LB_Week 
            Height          =   600
            Index           =   5
            Left            =   6000
            Top             =   630
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   1058
            BackColor       =   8421504
            Caption         =   "ã‡"
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
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D LB_Week 
            Height          =   600
            Index           =   6
            Left            =   7200
            Top             =   630
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   1058
            BackColor       =   16744576
            Caption         =   "ìy"
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
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   1
            Left            =   1200
            Top             =   1230
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   2
            Left            =   2400
            Top             =   1230
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   3
            Left            =   3600
            Top             =   1230
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   4
            Left            =   4800
            Top             =   1230
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   5
            Left            =   6000
            Top             =   1230
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   6
            Left            =   7200
            Top             =   1230
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            ForeColor       =   16744576
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   8
            Left            =   1200
            Top             =   2370
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   9
            Left            =   2400
            Top             =   2370
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   10
            Left            =   3600
            Top             =   2370
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   11
            Left            =   4800
            Top             =   2370
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   12
            Left            =   6000
            Top             =   2370
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   13
            Left            =   7200
            Top             =   2370
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            ForeColor       =   16744576
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   15
            Left            =   1200
            Top             =   3510
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   16
            Left            =   2400
            Top             =   3510
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   17
            Left            =   3600
            Top             =   3510
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   18
            Left            =   4800
            Top             =   3510
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   19
            Left            =   6000
            Top             =   3510
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   20
            Left            =   7200
            Top             =   3510
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            ForeColor       =   16744576
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   22
            Left            =   1200
            Top             =   4650
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   23
            Left            =   2400
            Top             =   4650
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   24
            Left            =   3600
            Top             =   4650
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   25
            Left            =   4800
            Top             =   4650
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   26
            Left            =   6000
            Top             =   4650
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   27
            Left            =   7200
            Top             =   4650
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            ForeColor       =   16744576
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   29
            Left            =   1200
            Top             =   5790
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   30
            Left            =   2400
            Top             =   5790
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   31
            Left            =   3600
            Top             =   5790
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   32
            Left            =   4800
            Top             =   5790
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   33
            Left            =   6000
            Top             =   5790
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   34
            Left            =   7200
            Top             =   5790
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            ForeColor       =   16744576
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   36
            Left            =   1200
            Top             =   6930
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   0
            Left            =   600
            Top             =   1770
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   1
            Left            =   1800
            Top             =   1770
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   2
            Left            =   2400
            Top             =   1770
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   2
            Left            =   3000
            Top             =   1770
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   3
            Left            =   3600
            Top             =   1770
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   3
            Left            =   4200
            Top             =   1770
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   4
            Left            =   4800
            Top             =   1770
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   4
            Left            =   5400
            Top             =   1770
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   5
            Left            =   6000
            Top             =   1770
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   5
            Left            =   6600
            Top             =   1770
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   6
            Left            =   7200
            Top             =   1770
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   6
            Left            =   7800
            Top             =   1770
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   7
            Left            =   600
            Top             =   2910
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   8
            Left            =   1200
            Top             =   2910
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   8
            Left            =   1800
            Top             =   2910
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   9
            Left            =   2400
            Top             =   2910
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   9
            Left            =   3000
            Top             =   2910
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   10
            Left            =   3600
            Top             =   2910
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   10
            Left            =   4200
            Top             =   2910
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   11
            Left            =   4800
            Top             =   2910
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   11
            Left            =   5400
            Top             =   2910
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   12
            Left            =   6000
            Top             =   2910
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   12
            Left            =   6600
            Top             =   2910
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   13
            Left            =   7200
            Top             =   2910
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   13
            Left            =   7800
            Top             =   2910
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   14
            Left            =   600
            Top             =   4050
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   15
            Left            =   1200
            Top             =   4050
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   15
            Left            =   1800
            Top             =   4050
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   16
            Left            =   2400
            Top             =   4050
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   16
            Left            =   3000
            Top             =   4050
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   17
            Left            =   3600
            Top             =   4050
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   17
            Left            =   4200
            Top             =   4050
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   18
            Left            =   4800
            Top             =   4050
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   18
            Left            =   5400
            Top             =   4050
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   19
            Left            =   6000
            Top             =   4050
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   19
            Left            =   6600
            Top             =   4050
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   20
            Left            =   7200
            Top             =   4050
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   20
            Left            =   7800
            Top             =   4050
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   21
            Left            =   600
            Top             =   5190
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   22
            Left            =   1200
            Top             =   5190
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   22
            Left            =   1800
            Top             =   5190
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   23
            Left            =   2400
            Top             =   5190
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   23
            Left            =   3000
            Top             =   5190
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   24
            Left            =   3600
            Top             =   5190
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   24
            Left            =   4200
            Top             =   5190
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   25
            Left            =   4800
            Top             =   5190
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   25
            Left            =   5400
            Top             =   5190
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   26
            Left            =   6000
            Top             =   5190
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   26
            Left            =   6600
            Top             =   5190
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   27
            Left            =   7200
            Top             =   5190
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   27
            Left            =   7800
            Top             =   5190
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   28
            Left            =   600
            Top             =   6330
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   29
            Left            =   1200
            Top             =   6330
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   29
            Left            =   1800
            Top             =   6330
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   30
            Left            =   2400
            Top             =   6330
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   30
            Left            =   3000
            Top             =   6330
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   31
            Left            =   3600
            Top             =   6330
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   31
            Left            =   4200
            Top             =   6330
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   32
            Left            =   4800
            Top             =   6330
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   32
            Left            =   5400
            Top             =   6330
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   33
            Left            =   6000
            Top             =   6330
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   33
            Left            =   6600
            Top             =   6330
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   34
            Left            =   7200
            Top             =   6330
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   34
            Left            =   7800
            Top             =   6330
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   35
            Left            =   600
            Top             =   7470
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   36
            Left            =   1200
            Top             =   7470
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Yakin 
            Height          =   600
            Index           =   36
            Left            =   1800
            Top             =   7470
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   0
            Left            =   0
            Top             =   1770
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   0
            Left            =   0
            Top             =   1230
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            ForeColor       =   255
            Caption         =   "1"
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D LB_Week 
            Height          =   600
            Index           =   0
            Left            =   0
            Top             =   630
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   1058
            BackColor       =   255
            Caption         =   "ì˙"
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
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   7
            Left            =   0
            Top             =   2370
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            ForeColor       =   255
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   14
            Left            =   0
            Top             =   3510
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            ForeColor       =   255
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   21
            Left            =   0
            Top             =   4650
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            ForeColor       =   255
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   28
            Left            =   0
            Top             =   5790
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            ForeColor       =   255
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D B1lb_DD 
            Height          =   540
            Index           =   35
            Left            =   0
            Top             =   6930
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   953
            BackColor       =   14737632
            ForeColor       =   255
            Caption         =   ""
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   7
            Left            =   0
            Top             =   2910
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   14
            Left            =   0
            Top             =   4050
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   21
            Left            =   0
            Top             =   5190
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   28
            Left            =   0
            Top             =   6330
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_Hiru 
            Height          =   600
            Index           =   35
            Left            =   0
            Top             =   7470
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            ForeColor       =   255
            Caption         =   "Çm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D lb_YYMM 
            Height          =   600
            Left            =   0
            Top             =   0
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   1058
            BackColor       =   16761087
            ForeColor       =   16711680
            Caption         =   "1998îN04åé"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr Çoñæí©"
               Size            =   20.25
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
      End
   End
   Begin VB.PictureBox H1_Area1 
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   120
      ScaleHeight     =   1290
      ScaleWidth      =   6315
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   945
      Width           =   6375
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   345
         Left            =   60
         Top             =   885
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   609
         ForeColor       =   12582912
         Caption         =   "     Ñ§Å®  éÊà¯êÊÅFãÛîíÇÕé©é–ÉJÉåÉìÉ_Ç≈Ç∑"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr Çoñæí©"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   1
      End
      Begin Cis3D_v60.CIS3D B1_Area22 
         Height          =   810
         Left            =   60
         Top             =   60
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   1429
         ForeColor       =   16711680
         Caption         =   "éÊà¯êÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cPositionX      =   270
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   75
            TabIndex        =   1
            Top             =   420
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
            cAutoEject      =   0   'False
         End
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1110
            Top             =   45
            Width           =   2490
            _ExtentX        =   4392
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
         Begin Cis3D_v60.CIS3D H1lb_Kojyonm 
            Height          =   360
            Left            =   1110
            Top             =   420
            Width           =   2490
            _ExtentX        =   4392
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
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   810
         Index           =   1
         Left            =   5565
         Top             =   60
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1429
         ForeColor       =   16711680
         Caption         =   "îN"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_YY 
            Height          =   375
            Left            =   30
            TabIndex        =   4
            Top             =   405
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            cFormat         =   "0000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "1998"
            MaxLength       =   4
            cDataType       =   1
            cDataReplace    =   1
            cFaZero         =   0
            cFbComma        =   0
            cGFormat        =   "0000"
            cILength        =   4
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D35 
         Height          =   810
         Left            =   3750
         Top             =   60
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1429
         ForeColor       =   16711680
         Caption         =   "éÊà¯êÊãÊï™"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         cPositionY      =   60
         Begin Cis3D_v60.CIS3D Torkb_PNLG 
            Height          =   360
            Left            =   30
            Top             =   405
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   635
            BackColor       =   12640511
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   9
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
            Begin VB.OptionButton H1Op_Torkb 
               BackColor       =   &H00C0E0FF&
               Caption         =   "éÛíç"
               Height          =   240
               Index           =   0
               Left            =   45
               TabIndex        =   2
               Top             =   75
               Value           =   -1  'True
               Width           =   825
            End
            Begin VB.OptionButton H1Op_Torkb 
               BackColor       =   &H00C0E0FF&
               Caption         =   "î≠íç"
               Height          =   240
               Index           =   1
               Left            =   900
               TabIndex        =   3
               Top             =   75
               Width           =   825
            End
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
      Left            =   7365
      ScaleHeight     =   660
      ScaleWidth      =   7740
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   945
      Width           =   7800
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   5160
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
         Left            =   6015
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
         Left            =   6870
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
      Begin CisBtn_60.CisBtn PB_NEXT 
         Height          =   585
         Left            =   3450
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
         Left            =   2595
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   1740
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
         Left            =   30
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
         Left            =   885
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
         Left            =   4305
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
   End
   Begin VB.Timer Timer1 
      Left            =   195
      Top             =   60
   End
   Begin VB.Timer Timer2 
      Left            =   615
      Top             =   60
   End
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      TabIndex        =   0
      Text            =   "Dummy"
      Top             =   30
      Width           =   240
   End
   Begin VB.PictureBox B1_Area1 
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8340
      Left            =   45
      ScaleHeight     =   8280
      ScaleWidth      =   6450
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2550
      Width           =   6510
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   600
         Left            =   60
         Top             =   90
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   1058
         BackColor       =   -2147483633
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'ìßñæ
            Caption         =   "ì¸"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   2
            Left            =   405
            TabIndex        =   12
            Top             =   195
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'ìßñæ
            Caption         =   "éÛ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   1
            Left            =   225
            TabIndex        =   11
            Top             =   15
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'ìßñæ
            Caption         =   "åé"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   0
            Left            =   45
            TabIndex        =   10
            Top             =   345
            Width           =   240
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   15
            X2              =   615
            Y1              =   0
            Y2              =   600
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   0
            X1              =   0
            X2              =   600
            Y1              =   0
            Y2              =   600
         End
      End
      Begin vsFlexLib.vsFlexArray vsGrid1 
         Height          =   8070
         Left            =   60
         TabIndex        =   5
         Top             =   90
         Width           =   6315
         _Version        =   196608
         _ExtentX        =   11139
         _ExtentY        =   14235
         _StockProps     =   228
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ConvInfo        =   1418783674
         FormatString    =   "    |    |    |    |    |    |    |    |    |    |    |    |    |"
         Rows            =   13
         Cols            =   14
         BackColor       =   12648447
         ForeColorFixed  =   16711680
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   600
         VirtualData     =   0   'False
         MultiTotals     =   0   'False
         BackColorAlternate=   12648447
         MergeCompare    =   1
      End
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
      Caption         =   "Åy ÉJÉåÉìÉ_É}ÉXÉ^Å@ÉÅÉìÉeÉiÉìÉX  Åz"
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
         Left            =   10620
         ScaleHeight     =   345
         ScaleWidth      =   3675
         TabIndex        =   7
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
            BackColor       =   16777088
            ForeColor       =   0
            Caption         =   "èCê≥"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14370
         Top             =   90
         Width           =   885
         _ExtentX        =   1561
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
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   285
         Left            =   8070
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
End
Attribute VB_Name = "CXM0040"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :   ÉJÉåÉìÉ_Å[É}ÉXÉ^  ÉÅÉìÉeÉiÉìÉX
'**       ÉtÉHÅ[ÉÄID    :   CXM0040
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :   2003/12/11  By CIS
'**       ïœ  çX  ì˙    :   2005.02.14  By CIS  éÊà¯êÊãÊï™èâä˙ï\é¶ÉpÉâÉÅÅ[É^í«â¡
'**                                             ( ÇPåÖñ⁄ÅF0/éÛíçÅA1/î≠íç )
'**       ïœ  çX  ì˙    :   2005.02.28  By CIS  â“ìÆãÊï™Ç…Åuó’éûÅví«â¡
'**       ïœ  çX  ì˙    :   2005.03.10  By CIS  â“ìÆãÊï™Åuó’éûÅvÇÃéûÅAï∂éöêFÇóŒÇ…
'**       ïœ  çX  ì˙    :   2006/03/28  By CIS  î≠íçÇ≈é©é–ÇÃèÍçáÇ‡éÛì¸íPà Ç≈ÇÃì¸óÕÇâ¬î\Ç…ïœçX
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    Dim wObj            As Object       ' µÃﬁºﬁ™∏ƒ‹∞∏¥ÿ±
    
    Dim ProcHB          As String * 2   ' ì¸óÕàÊêÿë÷
    Dim SyoriKB         As String       ' ì¸óÕàÊêÿë÷
    Dim SyoriNM         As String       ' ì¸óÕàÊêÿë÷
    Dim FormAct         As Boolean      ' Ã´∞—±∏√®Ãﬁ Ã◊∏ﬁ
    Dim SL_Select       As String       ' SQL SELECT
    Dim HD_Select       As String
    Dim cnt_ADD         As Long
    Dim Cnt_MNT         As Long
    Dim cnt_DEL         As Long
 
    Dim I               As Integer      ' ìYéö
    Dim J               As Integer      ' ìYéö
    Dim CL_Ukeire()     As String       ' ï\é¶ópéÛì¸ÉèÅ[ÉN
    Dim Ukeire_Su       As Integer      ' éÛì¸êîÉèÅ[ÉN
    Dim w_Kado          As Currency     ' â“ìÆì˙êîÉèÅ[ÉN
    Dim w_Yakin         As Currency     ' ñÈãŒì˙êîÉèÅ[ÉN
    Dim w_TKado         As Currency     ' Tâ“ìÆì˙êîÉèÅ[ÉN
    Dim Bd_SyoriKB      As String       ' É{ÉfÉBÅ[ópèàóùãÊï™
    Dim Save_Idx        As Integer      ' åªç›à íuÉCÉìÉfÉbÉNÉXÉèÅ[ÉN
    Dim Save_HYkb       As Byte         ' åªç›à íuíãñÈãÊï™
    Dim w_Cols          As Integer      ' èâä˙âÊñ ï\é¶óp(èâä˙íl 14)
    Dim wk_Col          As Integer      ' ïúãAóp
    Dim wk_Row          As Integer      '   Åh
    Dim st_Week         As Byte         ' äJénójì˙( 0:ì˙ójì˙, 1:åéójì˙ )
    
    Const c_Cols        As Integer = 11
    Const c_FYellow     As Long = &HC0FFFF      ' <îñÇ¢â©êF> ÉJÉìÉåÉìÉ_íãÅEñÈãŒÉpÉlÉã(ÉtÉHÅ[ÉJÉXéû)
    Const c_LFGray      As Long = &H8000000F    ' <îñÇ¢äDêF> ÉJÉìÉåÉìÉ_íãÅEñÈãŒÉpÉlÉã(ÉçÉXÉgÉtÉHÅ[ÉJÉXéû)
    Const c_SatDBlue    As Long = &HFF8080      ' <Ç‚Ç‚îñê¬> ÉJÉìÉåÉìÉ_ìyójì˙ÉoÉbÉNÉJÉâÅ[
    Const c_WeekEndD    As Long = &H808080      ' <îZÇ¢äDêF> ÉJÉìÉåÉìÉ_ójì˙ÉoÉbÉNÉJÉâÅ[(ìyì˙à»äO)
    Const c_CRed        As Long = &HFF&         ' <ñæÇÈÇ¢ê‘> ÉJÉìÉåÉìÉ_íãÅEñÈãŒï∂éöêFÅió’éûâ“ìÆà»äOéûÅj     '2005.03.10 Add
    Const c_CGreen      As Long = &HC000&       ' <ñæÇÈÇ¢óŒ> ÉJÉìÉåÉìÉ_íãÅEñÈãŒï∂éöêFÅió’éûâ“ìÆéûÅj         '2005.03.10 Add

    Dim mJisyaKJ        As String               ' é©é–çHèÍ 2006/03/28 Insert

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
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
'   #------------------#
'   #  çÄ ñ⁄ èâ ä˙ âª   #
'   #------------------#
    Call CisFrm.InitFld
    
'   *---------------*
'   *   ójì˙ÉZÉbÉg  *
'   *---------------*
    Call DayOfTheWeek_Set
    
    Call CalenClr
    iH1_YY = Format(Date, "YYYY")
    w_Cols = c_Cols

'+--------------------------------+
'+  ÉOÉäÉbÉhèâä˙ê›íË(éÊà¯êÊèÓïÒ)
'+--------------------------------+
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 600
        .PatanMax = 1
        .InitGet = False
        .Init
        .NumberSet = All
    End With
    vsGrid1.Cols = w_Cols
    Call GridInit
    Call CisVsGrid.ReNumber

'
    Dummy.Left = -1000
    Dummy.Enabled = False
'
    Timer1.Interval = 1000   ' éûçèï\é¶ä‘äu
    Timer1.Enabled = True    ' éûçèï\é¶äJén
    Timer2.Interval = 500    ' èàóùãÊï™ì_ñ≈ä‘äu
    Timer2.Enabled = False   ' èàóùãÊï™ì_ñ≈ã÷é~

'-------------------------------------------< 2005.02.14 MNT START >
'    H1Op_Torkb(0) = True
    gInt = CisFun.Val2(Mid$(Command, 1, 1))
    H1Op_Torkb(gInt) = True
'-------------------------------------------< 2005.02.14 MNT  END  >
    
'   #------------------#
'   #    èâä˙ì¸óÕê›íË   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  èàóùãÊï™èâä˙ï\é¶ #
'   #------------------#
    SyoriKB = "MNT"
    Call SyoriName(SyoriKB)

'   #-------------------#
'   # âÊñ íÜâõà íuï\é¶   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
     
    Me.Show
End Sub
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    ÅsÇ`ÇÉÇîÇâÇñÇÅÇîÇÖ ÉCÉxÉìÉgÅt              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    CisFrm.GrpFocGotColor = gIGotSel
    CisFrm.GrpFocLostColor = gILostSel
    
    cnt_ADD = 0
    Cnt_MNT = 0
    cnt_DEL = 0
'------------------------------------------------ 2006/03/28 Insert
' ÉVÉXÉeÉÄèÓïÒ(é©é–çHèÍ)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 7
    Call IRNRead("", 1)
    mJisyaKJ = Mid(IRN.oyian, 1, 1)
'------------------------------------------------ 2006/03/28 Insert End
    
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås ÇjÇÖÇôÇcÇèÇóÇéÉCÉxÉìÉg Åt               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås QueryUnloadÉCÉxÉìÉg Åt                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    Cancel = 1
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄÅs ÉtÉ@ÉìÉNÉVÉáÉìÉLÅ[ópÉAÉNÉZÉâÉåÅ[É^ Åt       +
'+----------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    If Not FormAct Then Exit Function
    
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    
    Select Case Key_Code
        Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click        'ÅyêVãKÅz
        Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click        'ÅyèCê≥Åz
        Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click        'ÅyçÌèúÅz
        Case vbKeyF6:       If PB_Auto.Visible Then Call PB_Auto_Click      'Åyé©ìÆçÏê¨Åz
        Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click      'ÅyåüçıÅz
        Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click      'Åyå„ëﬁÅz
        Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click      'ÅyëOêiÅz
        Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_Can_Click        'ÅyéÊè¡Åz
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click        'ÅyèIóπÅz
        Case vbKeyReturn
                            If PB_ENT.Visible Then                          'Åyì¸óÕÅz
                               Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                               Call PB_ENT_Click
                               Dummy.Enabled = False
                            End If
        
        Case vbKeySpace
                            If ProcHB = "B2" Then
                                If Save_HYkb = 0 Then
                                    If B1lb_DD(Save_Idx).Visible Then Call PB_Hiru_Click(Save_Idx)
                                    GoTo Key_Acc_ED
                                Else
                                    If B1lb_DD(Save_Idx).Visible Then Call PB_Yakin_Click(Save_Idx)
                                    GoTo Key_Acc_ED
                                End If
                            End If
        
'   < åéÅEéÛì¸íPà Ç≈ÇÃà⁄ìÆ(Body2) >
        Case vbKeyNumpad2, vbKeyNumpad4, vbKeyNumpad6, vbKeyNumpad8
                            If ProcHB = "B2" Then Call DataChange(Key_Code) Else Exit Function
'   < ÉJÅ[É\ÉãÇÃà⁄ìÆ(Body1) >
        Case vbKeyDown, vbKeyUp, vbKeyLeft, vbKeyRight
                            If ProcHB = "B2" Then Call Idou_Forcus(Key_Code) Else Exit Function
    
        Case Else:          Exit Function
    End Select
    
Key_Acc_ED:
    Key_Acc = True
End Function
'+----------------------------+
'+    êV  ãK  ÉL Å[(F1)       +
'+----------------------------+
Private Sub PB_ADD_Click()
    SyoriKB = "ADD"
    Call CalenClr
    Call SyoriName(SyoriKB)
    
    iH1_Torcd.SetFocus
End Sub
'+----------------------------+
'+    èC  ê≥  ÉL Å[(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    SyoriKB = "MNT"
    Call CalenClr
    Call SyoriName(SyoriKB)
    iH1_Torcd.SetFocus
End Sub
'+----------------------------+
'+    çÌ  èú  ÉL Å[(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If ProcHB = "H1" Then
        SyoriKB = "DEL"
        Call CalenClr
        Call SyoriName(SyoriKB)
        iH1_Torcd.SetFocus
    Else
        If Bd_SyoriKB = "DEL" Then
            Bd_SyoriKB = ""
            vsGrid1.ForeColor = vbBlack
            vsGrid1.BackColorSel = &H8000000D
        Else
            Bd_SyoriKB = "DEL"
            vsGrid1.ForeColor = vbRed
            vsGrid1.BackColorSel = &HC0C0FF
        End If
    End If
End Sub
'+----------------------------+
'+    ëO  ï≈  ÉL Å[(PageUp)   +
'+----------------------------+
Private Sub PB_Back_Click()

    Dummy.Enabled = True
    Dummy.SetFocus
    DoEvents
    
    Call CisFrm.MousePT(11)
    
    SL_Select = "SELECT DISTINCT éÊà¯êÊãÊï™, éÊà¯êÊ" & vbCrLf
    SL_Select = SL_Select & " FROM ÉJÉåÉìÉ_É}ÉXÉ^" & vbCrLf
    SL_Select = SL_Select & " WHERE  Cast( éÊà¯êÊãÊï™ As Varchar(1) ) + éÊà¯êÊ"
    SL_Select = SL_Select & " < '" & CisFun.Val2(H1Op_Torkb(0).Tag) & RTrim$(iH1_Torcd) & "'" & vbCrLf
    SL_Select = SL_Select & " AND    îN = '" & iH1_YY & "'" & vbCrLf
    SL_Select = SL_Select & " ORDER BY éÊà¯êÊãÊï™ DESC, éÊà¯êÊ DESC" & vbCrLf
    
    Call DBInput("Back")
    iH1_Torcd.SetFocus

Back_ED:
    Call CisFrm.MousePT(1)
    Dummy.Enabled = False
End Sub
'+----------------------------+
'+    éü  ï≈  ÉL Å[(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    
    Dummy.Enabled = True
    Dummy.SetFocus
    DoEvents
    
    Call CisFrm.MousePT(11)
    
    SL_Select = "SELECT DISTINCT éÊà¯êÊãÊï™, éÊà¯êÊ" & vbCrLf
    SL_Select = SL_Select & " FROM ÉJÉåÉìÉ_É}ÉXÉ^" & vbCrLf
    SL_Select = SL_Select & " WHERE  Cast( éÊà¯êÊãÊï™ As Varchar(1) ) + éÊà¯êÊ"
    SL_Select = SL_Select & " > '" & CisFun.Val2(H1Op_Torkb(0).Tag) & RTrim$(iH1_Torcd) & "'" & vbCrLf
    SL_Select = SL_Select & " AND    îN = '" & iH1_YY & "'" & vbCrLf
    SL_Select = SL_Select & " ORDER BY éÊà¯êÊãÊï™, éÊà¯êÊ" & vbCrLf
    
    Call DBInput("Next")
    iH1_Torcd.SetFocus

Next_ED:
    Call CisFrm.MousePT(1)
    Dummy.Enabled = False
End Sub
'+---------------------+
'+  ÉJÉåÉìÉ_é©ìÆçÏê¨     +
'+---------------------+
Private Sub PB_Auto_Click()
    CXM0041.Show vbModal
    Unload CXM0041
    Set CXM0041 = Nothing
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
'+    éÊ  è¡  ÉL Å[(Escape)   +
'+----------------------------+
Private Sub PB_Can_Click()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "B1" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        
        w_Cols = c_Cols
        vsGrid1.Cols = w_Cols
        Call GridInit
        Call CisVsGrid.ReNumber
        
        Bd_SyoriKB = ""
        vsGrid1.ForeColor = &HF
        If SyoriKB = "Lock" Then Call CalenClr
        GoTo PB_CAN_Ed
    End If
    
    If ProcHB = "B2" Then
        ProcHB = "B1"
        Call DispChange(ProcHB)
        Call CalenClr
        GoTo PB_CAN_Ed
    End If

PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    éÊà¯êÊåüçı              +
'+----------------------------+
Private Sub PB_Look_Click()

    RV_Left = 0:    RV_Top = 0
    Select Case CisFun.Val2(H1Op_Torkb(0).Tag)
        Case 0:     RV_TorKb = 3
        Case Else:  RV_TorKb = 8
    End Select
    
    CKK0025.Show vbModal

    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RTrim$(RV_TorcdK)
        Call iH1_Torcd_LostFocus
    End If
    
    iH1_Torcd.SetFocus
End Sub
'*--------------------------------------*
'*      éÊà¯êÊ      ( iH1_Torcd )       *
'*--------------------------------------*
Private Sub iH1_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    
    Call GetTorcd
End Sub
'+---------------------+
'+   H1Op_Torkb
'+---------------------+
Private Sub H1Op_Torkb_Click(Index As Integer)
    H1Op_Torkb(0).Tag = Index
End Sub
Private Sub H1Op_Torkb_GotFocus(Index As Integer)
    CisFrm.GrpFocus H1Op_Torkb
End Sub
Private Sub H1Op_Torkb_LostFocus(Index As Integer)
    CisFrm.GrpFocus H1Op_Torkb, Lost
End Sub
'*---------------------------------*
'*   iH1_YY         ëŒè€îN          *
'*---------------------------------*
Private Sub iH1_YY_LostFocus()
    If CisFun.Val2(iH1_YY) < 1000 Then
        iH1_YY = 2000 + CisFun.Val2(iH1_YY)
    End If
End Sub
'+-------------------------------------+
'+          ójÅ@ì˙Å@ÉpÅ@ÉlÅ@Éã          +
'+-------------------------------------+
Private Sub LB_Week_Click(Index As Integer)
    Static wWeek    As Integer
    Static sw_KB    As Integer
    
    sw_KB = 0
    For wWeek = Index To 36 Step 7
        If B1lb_DD(wWeek).Visible Then
            Call Set_KadoPnl(wWeek)                 ' < ÉpÉlÉãï\é¶ì‡óeÉZÉbÉg >
            If sw_KB = 0 Then sw_KB = wWeek
        End If
    Next wWeek

    Call Set_KadoSu                                 ' < â“ìÆì˙êî >
    
    PB_Hiru(Save_Idx).BackColor = c_LFGray
    PB_Yakin(Save_Idx).BackColor = c_LFGray
    
    Save_Idx = sw_KB
    Save_HYkb = 0
    PB_Hiru(Save_Idx).BackColor = c_FYellow

End Sub
'+-----------------------------------------+
'+          ì˙Å@Ç…Å@ÇøÅ@ÉpÅ@ÉlÅ@Éã          +
'+-----------------------------------------+
Private Sub B1lb_DD_Click(Index As Integer)
    Call Set_KadoPnl(Index)                         ' < ÉpÉlÉãï\é¶ì‡óeÉZÉbÉg >
    Call Set_KadoSu                                 ' < â“ìÆì˙êî >
End Sub
'+-------------------------------------+
'+          íãÅ@ãŒÅ@ÉpÅ@ÉlÅ@Éã          +
'+-------------------------------------+
Private Sub PB_Hiru_Click(Index As Integer)
    Call Set_KadoPnl(Index, "Hiru")                 ' < ÉpÉlÉãï\é¶ì‡óeÉZÉbÉg >
    Call Set_KadoSu                                 ' < â“ìÆì˙êî >
End Sub
'+-------------------------------------+
'+          ñÈÅ@ãŒÅ@ÉpÅ@ÉlÅ@Éã          +
'+-------------------------------------+
Private Sub PB_Yakin_Click(Index As Integer)
    Call Set_KadoPnl(Index, "Yakin")                ' < ÉpÉlÉãï\é¶ì‡óeÉZÉbÉg >
    Call Set_KadoSu                                 ' < â“ìÆì˙êî >
End Sub
'*-----------------------------*
'*      â“ ì≠ ì˙ Ép Él Éã       *
'*-----------------------------*
Private Sub Set_KadoPnl(Index As Integer, Optional pKB As String = "")
    
    PB_Hiru(Save_Idx).BackColor = c_LFGray
    PB_Yakin(Save_Idx).BackColor = c_LFGray
    
'   < ójì˙ÅEì˙Ç…ÇøÅEíããŒ ÉNÉäÉbÉN >
    If pKB <> "Yakin" Then
''2005.02.28 Upd Start
''        If Trim$(PB_Hiru(Index)) = "" Then
''            PB_Hiru(Index) = "ãx"
''        Else
''            PB_Hiru(Index) = ""
''        End If
        Select Case Trim$(PB_Hiru(Index))
            Case ""
                PB_Hiru(Index) = "ó’"
                PB_Hiru(Index).ForeColor = c_CGreen             '2005.03.10 Add
            Case "ó’"
                PB_Hiru(Index) = "ãx"
                PB_Hiru(Index).ForeColor = c_CRed               '2005.03.10 Add
            Case Else
                PB_Hiru(Index) = ""
                
        End Select
''2005.02.28 Upd End
    End If
    
'   < ójì˙ÅEì˙Ç…ÇøÅEñÈãŒ ÉNÉäÉbÉN >
    If pKB <> "Hiru" Then
''2005.02.28 Upd Start
''        If Trim$(PB_Yakin(Index)) = "" Then
''            PB_Yakin(Index) = "ãx"
''        Else
''            PB_Yakin(Index) = ""
''        End If
        Select Case Trim$(PB_Yakin(Index))
            Case ""
                PB_Yakin(Index) = "ó’"
                PB_Yakin(Index).ForeColor = c_CGreen            '2005.03.10 Add
            Case "ó’"
                PB_Yakin(Index) = "ãx"
                PB_Yakin(Index).ForeColor = c_CRed              '2005.03.10 Add
            Case Else
                PB_Yakin(Index) = ""
        End Select
''2005.02.28 Upd End
    End If

    Save_Idx = Index
    If pKB = "Yakin" Then
        Save_HYkb = 1
        PB_Yakin(Save_Idx).BackColor = c_FYellow
    Else
        Save_HYkb = 0
        PB_Hiru(Save_Idx).BackColor = c_FYellow
    End If
End Sub
'*-----------------------------*
'*      â“ ìÆ ì˙ êî éÊ ìæ       *
'*-----------------------------*
Private Sub Set_KadoSu()
    w_Kado = 0: w_Yakin = 0: w_TKado = 0

    For I = 0 To 36
        If B1lb_DD(I).Visible Then
            If PB_Hiru(I) <> "ãx" And PB_Hiru(I) <> "ó’" Then w_Kado = w_Kado + 1           '2005.02.28 Upd
            If PB_Yakin(I) <> "ãx" And PB_Yakin(I) <> "ó’" Then w_Yakin = w_Yakin + 1       '2005.02.28 Upd
        End If
    Next I
    w_TKado = w_Kado + w_Yakin
    w_TKado = w_TKado / 2

    B1lb_Nisu(0) = Format(w_Kado, "#0")
    B1lb_Nisu(1) = Format(w_Yakin, "#0")
    B1lb_Nisu(2) = Format(w_TKado, "#0.0")
End Sub
'**************************************
'*                                    *
'*      èàóùãÊï™Ç…ÇÊÇËãÊï™ñºï\é¶        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   
    Dim w_Color     As Long
    
    w_Color = &H80000008 'èâä˙êFæØƒ(çï)
    Select Case SyoriKB
        Case "ADD"
                    SyoriNM = "í«â¡"
                    K_Sykbnm.ForeColor = &HFFFF&   '(â©êF)
        Case "MNT"
                    SyoriNM = "èCê≥"
                    K_Sykbnm.ForeColor = &HFF0000  '(ê¬êF)
        Case "DEL"
                    SyoriNM = "çÌèú"
                    K_Sykbnm.ForeColor = vbRed     '(ê‘êF)
                    w_Color = vbRed
        Case "Lock"
                    SyoriNM = "è∆âÔ"
                    K_Sykbnm.ForeColor = &HFF00FF  '(ÉsÉìÉN)
                    w_Color = &HFF00FF
        Case Else
                    SyoriNM = ""
    End Select
    
    For Each wObj In Me.Controls
        If wObj.Name Like "i*" Then wObj.ForeColor = w_Color
        If wObj.Name Like "*Comb*" Then wObj.ForeColor = w_Color
    '      * PF-Key ï\é¶ê›íË *
        If ProcHB = "H1" Then
            If wObj.Name = "PB_ADD" Or wObj.Name = "PB_MNT" Or _
               wObj.Name = "PB_DEL" Then
                    wObj.Visible = Not (UCase(Mid(wObj.Name, 4, 3)) = SyoriKB)
            End If
        End If
    Next wObj
     
    If SyoriKB = "DEL" Then
        vsGrid1.ForeColor = vbRed
        vsGrid1.BackColor = &HC0C0FF
    Else
        vsGrid1.ForeColor = &H80000008
        vsGrid1.BackColor = &HC0FFFF
    End If
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

    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
       Case "H1":       H1Mode = True:  H1Color = gPGotSel
       Case "B1":       B1Mode = True:  B1Color = gPGotSel
       Case "B2":       B2Mode = True:  B2Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B2_Area1.Enabled = B2Mode
    B2_Area1.BackColor = B2Color
    P_TenKey.Visible = B2Mode
    
    PB_ADD.Visible = H1Mode
    PB_MNT.Visible = H1Mode
    PB_DEL.Visible = H1Mode
    
    If ProcHB = "B1" And SyoriKB <> "Lock" Then PB_DEL.Visible = B1Mode

    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    PB_Auto.Visible = H1Mode
    
    If ProcHB = "H1" Then
        vsGrid1.ForeColor = vbBlack
        vsGrid1.BackColorSel = &H8000000D
        vsGrid1.HighLight = flexHighlightNever
    Else
        vsGrid1.HighLight = flexHighlightAlways
    End If
' ¿ﬁ–∞çÄñ⁄(ì¸óÕã÷é~)
    Dummy.Enabled = False
'
End Function
'****************************
'*      ì¸óÕèàóù(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
' << çÌèúèàóù >>
        If SyoriKB = "DEL" Then
            If SyoriKB = "DEL" Then
                If Not DBDelete Then GoTo ReturnPress_Ed
                w_Cols = c_Cols
                Call GridInit
                CisVsGrid.Clear
                ProcHB = "H1"
            End If
        Else
            ProcHB = "B1"
        End If
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        GoTo ReturnPress_Ed
    End If
   
'   < â“ì≠ì˙ÉOÉäÉbÉh >
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        
        If Bd_SyoriKB = "DEL" Then
            wk_Col = vsGrid1.Col: wk_Row = vsGrid1.Row
            If Not Uke_DBDelete Then GoTo ReturnPress_Ed
            Call GridReDisp                                 'GRIDçƒï\é¶
            vsGrid1.Col = wk_Col: vsGrid1.Row = wk_Row
            GoTo ReturnPress_Ed
        End If
        
        Call Body2Set
        
        ProcHB = "B2"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
    
'   < ÉJÉåÉìÉ_ÉpÉlÉã >
    If ProcHB = "B2" Then
        If Not DBPut Then GoTo ReturnPress_Ed
        
        wk_Col = vsGrid1.Col: wk_Row = vsGrid1.Row
        Call CalenClr
        Call GridReDisp                                 'GRIDçƒï\é¶
        vsGrid1.Col = wk_Col: vsGrid1.Row = wk_Row
        ProcHB = "B1"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    
    End If
ReturnPress_Ed:
    Call CLMClose
    Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1¡™Ø∏èàóù
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
 
    If Trim(iH1_YY) = "" Then iH1_YY = Format(Date, "YYYY")

' éÊà¯êÊÉ}ÉXÉ^É`ÉFÉbÉN
    If SyoriKB <> "DEL" Then
        If Not GetTorcd Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    éÊà¯êÊÉ}ÉXÉ^ñ¢ìoò^            "
                .MB_Button = Error
                .MBOX
            End With
            Exit Function
        End If
    End If
    
    If DBInput("INV") Then
        If SyoriKB = "DEL" Then
            Head1Chk = True
            Exit Function
        End If
        If SyoriKB = "ADD" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(1) = "          ìoò^çœÇ≈Ç∑ÅB            "
                .MB_MSG(3) = "    çƒìxíäèoèåèÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢            "
                .MB_Button = Error
                .MBOX
            End With
            Exit Function
        End If
    Else
        If SyoriKB = "DEL" Or SyoriKB = "MNT" Or SyoriKB = "Lock" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(1) = "    äYìñÉfÅ[É^Ç™ë∂ç›ÇµÇ‹ÇπÇÒ°            "
                .MB_MSG(3) = "    çƒìxíäèoèåèÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢            "
                .MB_Button = Error
                .MBOX
            End With
            Exit Function
        End If
    End If
    Head1Chk = True

End Function
'****************************
'*      BODY1¡™Ø∏èàóù
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False

    With vsGrid1

'       < Row & RowSelÅACol & ColSel >
        If .Row > .RowSel Then gInt = .Row: .Row = .RowSel: .RowSel = gInt
        If .Col > .ColSel Then gInt = .Col: .Col = .ColSel: .ColSel = gInt

'       < Select Check >
        If .Col > Ukeire_Su Then
            .Col = Ukeire_Su
            .ColSel = Ukeire_Su
            Exit Function
        End If
        If .ColSel > Ukeire_Su Then .ColSel = Ukeire_Su
        
        If Bd_SyoriKB <> "DEL" Then
            If .Row <> .RowSel Then .RowSel = .Row: Exit Function
            If .Col > Ukeire_Su Then Exit Function
            If .ColSel > Ukeire_Su Then .ColSel = Ukeire_Su
        Else
            gStr = ""
            For I = .Row To .RowSel
                For J = .Col To .ColSel
                    If RTrim$(.TextMatrix(I, J)) <> "" Then gStr = "X": Exit For
                Next J
                If RTrim$(gStr) <> "" Then Exit For
            Next I
        
            If RTrim$(gStr) = "" Then Exit Function
        End If
    End With

Body1Chk_ED:
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
    
    Call CalenClr
    
    If InputKey <> "INV" Then
        If Not CLMRead(SL_Select, 1) Then
            If InputKey = "Next" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    ç≈èIÉåÉRÅ[ÉhÇ≈Ç∑            "
                    .MB_Title = "Åyéüï≈Åz"
                    .MB_Button = OK
                    .MBOX
                End With
            Else
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    êÊì™ÉåÉRÅ[ÉhÇ≈Ç∑            "
                    .MB_Title = "ÅyëOï≈Åz"
                    .MB_Button = OK
                    .MBOX
                End With
            End If
            GoTo DBInput_Ed
        Else
            iH1_Torcd = RTrim$(CLM.éÊà¯êÊ)
            H1Op_Torkb(CLM.éÊà¯êÊãÊï™) = True
            H1Op_Torkb(0).Tag = CLM.éÊà¯êÊãÊï™
        End If
    End If
    
    SL_Select = "SELECT CL.*, TR.ó™èÃ, TR.çHèÍñº FROM ÉJÉåÉìÉ_É}ÉXÉ^ CL" & vbCrLf
    SL_Select = SL_Select & " Left Outer Join éÊà¯êÊÉ}ÉXÉ^ TR" & vbCrLf
    SL_Select = SL_Select & " On    TR.éÊà¯êÊCD = CL.éÊà¯êÊ" & vbCrLf
    SL_Select = SL_Select & " And   TR.éÊà¯êÊãÊï™ = CL.éÊà¯êÊãÊï™" & vbCrLf
    SL_Select = SL_Select & " WHERE CL.éÊà¯êÊ     = " & "'" & RTrim$(iH1_Torcd) & "'" & vbCrLf
    SL_Select = SL_Select & " And   CL.éÊà¯êÊãÊï™ = " & CisFun.Val2(H1Op_Torkb(0).Tag) & vbCrLf
    SL_Select = SL_Select & " And   CL.îN         = " & "'" & RTrim$(iH1_YY) & "'" & vbCrLf
    SL_Select = SL_Select & " And   CL.åé Between '01' And '12' " & vbCrLf
    SL_Select = SL_Select & " ORDER BY CL.éÊà¯êÊ, CL.îN, CL.éÛì¸, CL.åé" & vbCrLf
    
    HD_Select = SL_Select
    
    If Not CLMRead(SL_Select) Then
        Call GridHead_Set
        CisVsGrid.ReNumber
    Else
        iH1_Torcd = RTrim$(CLM.éÊà¯êÊ)
        H1lb_Tornm = RTrim$(CLM.ó™èÃ)
        H1lb_Kojyonm = RTrim$(CLM.çHèÍñº)
        H1Op_Torkb(CLM.éÊà¯êÊãÊï™) = True
        H1Op_Torkb(0).Tag = CLM.éÊà¯êÊãÊï™
        If RTrim$(CLM.éÊà¯êÊ) = "" Then
            H1lb_Tornm = ""
            H1lb_Kojyonm = "ÅôÅ@é©é–ÉJÉåÉìÉ_Å@Åô"
        Else
            If RTrim$(CLM.ó™èÃ) = "" Then
                H1lb_Tornm = "Å¶Å@ñ¢Å@ìoÅ@ò^Å@Å¶"
                H1lb_Kojyonm = ""
            End If
        End If
        
        Call GridHead_Set(InputKey)
        CisVsGrid.ReNumber
        Call GridSet
        
        DBInput = True
    End If
    
    vsGrid1.Row = Format(Now(), "m"): vsGrid1.Col = 1
DBInput_Ed:
    Call CLMClose           ' åãâ æØƒ∏€∞Ωﬁ
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'*******************************
'*      É^ÉCÉgÉãÉoÅ[ï“èW        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID
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
'******************************************'
'***** ÉJÉåÉìÉ_Å[èÓïÒÉNÉäÉA            *****
'******************************************'
Private Sub CalenClr()

    lb_YYMM = ""
    For Each wObj In PB_Hiru
        wObj.Caption = ""
    Next wObj
    For Each wObj In PB_Yakin
        wObj.Caption = ""
    Next wObj
    For Each wObj In B1lb_DD
        wObj.Caption = ""
    Next wObj
    LB_Week(0).Tag = ""
    
    B1lb_Nisu(0).Caption = ""
    B1lb_Nisu(1).Caption = ""
    B1lb_Nisu(2).Caption = ""
    
    PB_Hiru(Save_Idx).BackColor = c_LFGray
    PB_Yakin(Save_Idx).BackColor = c_LFGray
    Save_Idx = 0
    Save_HYkb = 0
    
End Sub
'**************************************
'*                                    *
'*          ÉfÅ[É^ÉxÅ[ÉXçXêV           *
'*                                    *
'**************************************
'+-------------------------------------+
'+      çÌ èú èà óù ( Head Lebel )     +
'+-------------------------------------+
Private Function DBDelete()
    DBDelete = False

'  ( ämîFÉÅÉbÉZÅ[ÉW )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 6
        .MB_MSG(1) = "    Åy çÌèúèàóùämîF Åz            "
        .MB_MSG(3) = "    éÊà¯êÊ ÅÅ  " & Trim(iH1_Torcd)
        .MB_MSG(4) = "    éÊà¯êÊãÊï™ ÅÅ  " & H1Op_Torkb(CisFun.Val2(H1Op_Torkb(0).Tag)).Caption
        .MB_MSG(5) = "    îN ÅÅ  " & Trim(iH1_YY)
        .MB_Title = "çÌèúèàóù"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
 
    CisDB.DBTran TransBegin
    
' ( çÌèúèàóù:ÉJÉåÉìÉ_É}ÉXÉ^ )
    SL_Select = "delete FROM ÉJÉåÉìÉ_É}ÉXÉ^"
    SL_Select = SL_Select & " WHERE éÊà¯êÊ = " & "'" & Trim(iH1_Torcd) & "'"
    SL_Select = SL_Select & " And   éÊà¯êÊãÊï™ = " & CisFun.Val2(H1Op_Torkb(0).Tag)
    SL_Select = SL_Select & " And   îN        = " & "'" & Trim(iH1_YY) & "'"
    
    On Error GoTo DBDelete_Err
    
    With CisDB
        .SQL = SL_Select
        .DBExec
    
        cnt_DEL = cnt_DEL + .RecordCount
    End With

    On Error GoTo 0
 
    CisDB.DBTran TransCommit
    
    iT1_Delsu = Format(cnt_DEL, "#,###")
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
    Dim Up_UkeireS     As String
    Dim Up_UkeireE     As String
    Dim Up_YYMM        As String
    DBPut = False

'  ( ëIëîÕàÕäiî[ )
    With vsGrid1
        Up_UkeireS = RTrim$(.TextMatrix(0, .Col))
        Up_UkeireE = RTrim$(.TextMatrix(0, .ColSel))
        Up_YYMM = Trim(iH1_YY) & Format(.Row, "00")
    End With
    
'  ( ämîFÉÅÉbÉZÅ[ÉW )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = " éÊà¯êÊ ÅÅ " & Trim(iH1_Torcd)
        If Up_UkeireS <> Up_UkeireE Then
            .MB_MSG(3) = " éÛì¸ ÅÅ " & Trim(Up_UkeireS) & " Å` " & Trim(Up_UkeireE)
        Else
            .MB_MSG(3) = " éÛì¸ ÅÅ " & Trim(Up_UkeireS)
        End If
        .MB_MSG(4) = " îNåé ÅÅ " & Trim(Up_YYMM)
        
        Select Case Bd_SyoriKB
            Case "ADD":         .MB_Title = "í«â¡èàóù":     .MB_MSG(1) = "   Åy í«â¡èàóùämîF Åz      "
            Case "MNT":         .MB_Title = "èCê≥èàóù":     .MB_MSG(1) = "   Åy èCê≥èàóùämîF Åz      "
            Case Else:          .MB_Title = "àÍäáèàóù":     .MB_MSG(1) = "   Åy àÍäáèàóùämîF Åz      "
        End Select
        
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
'   ( ññì˙ )
    CisFun.Date_RsFormat = fDD
    CisFun.Date_End RTrim$(iH1_YY) & Format(vsGrid1.Row, "00") & "01"
    
'< ÉèÅ[ÉNÉZÉbÉg >
    Call ItemsClearCLM
    With CLM
        .éÊà¯êÊãÊï™ = CisFun.Val2(H1Op_Torkb(0).Tag)
        .éÊà¯êÊ = Trim(iH1_Torcd)
        .îN = Trim(iH1_YY)
        .åé = Format(vsGrid1.TextMatrix(vsGrid1.Row, 0), "00")
        .ññì˙ = CisFun.Date_Result
        .â“ìÆì˙êî = CisFun.Val2(B1lb_Nisu(2))
        .íããŒì˙êî = CisFun.Val2(B1lb_Nisu(0))
        .ñÈãŒì˙êî = CisFun.Val2(B1lb_Nisu(1))

        .â“ìÆãÊï™ = Space(31)
        .ñÈãŒãÊï™ = Space(31)
        .ó’éûíããÊï™ = Space(31)                 '2005.02.28 Add
        .ó’éûñÈãÊï™ = Space(31)                 '2005.02.28 Add
        
        For I = 0 To 36
            If B1lb_DD(I).Visible Then
                If RTrim$(PB_Hiru(I).Caption) = "" Then
                    Mid(.â“ìÆãÊï™, CisFun.Val2(B1lb_DD(I)), 1) = "0"
                Else
                    Mid(.â“ìÆãÊï™, CisFun.Val2(B1lb_DD(I)), 1) = "1"
                End If
            
                If RTrim$(PB_Yakin(I).Caption) = "" Then
                    Mid(.ñÈãŒãÊï™, CisFun.Val2(B1lb_DD(I)), 1) = "0"
                Else
                    Mid(.ñÈãŒãÊï™, CisFun.Val2(B1lb_DD(I)), 1) = "1"
                End If
''2005.02.28 Add Start
                If RTrim$(PB_Hiru(I).Caption) = "ó’" Then
                    Mid(.ó’éûíããÊï™, CisFun.Val2(B1lb_DD(I)), 1) = "2"
                Else
                    Mid(.ó’éûíããÊï™, CisFun.Val2(B1lb_DD(I)), 1) = "0"
                End If
                If RTrim$(PB_Yakin(I).Caption) = "ó’" Then
                    Mid(.ó’éûñÈãÊï™, CisFun.Val2(B1lb_DD(I)), 1) = "2"
                Else
                    Mid(.ó’éûñÈãÊï™, CisFun.Val2(B1lb_DD(I)), 1) = "0"
                End If
''2005.02.28 Add End
            End If
        Next I
    
    End With
    
    CisDB.DBTran TransBegin
    On Error GoTo DBPut_Err
    
    With vsGrid1
        For I = .Col To .ColSel
            CLM.éÛì¸ = RTrim$(.TextMatrix(0, I))
        
            If RTrim$(.TextMatrix(.Row, I)) = "" Then
                Call CLMInsert
                cnt_ADD = cnt_ADD + 1
            Else
                Call CLMUpdate
                Cnt_MNT = Cnt_MNT + 1
            End If
        Next I
    End With
    
    On Error GoTo 0
    CisDB.DBTran TransCommit
    
    iT1_Addsu = Format(cnt_ADD, "#,###")
    iT1_Mntsu = Format(Cnt_MNT, "#,###")
    DBPut = True
    Exit Function

DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+-------------------------------------+
'+      çÌ èú èà óù ( éÛì¸ÅEîNåé )      +
'+-------------------------------------+
Private Function Uke_DBDelete()
    Dim Del_UkeireS     As String
    Dim Del_UkeireE     As String
    Dim Del_YYMMS       As String
    Dim Del_YYMME       As String
    
    Uke_DBDelete = False

'  ( ëIëîÕàÕäiî[ )
    With vsGrid1
        Del_UkeireS = RTrim$(.TextMatrix(0, .Col))
        Del_UkeireE = RTrim$(.TextMatrix(0, .ColSel))
        Del_YYMMS = RTrim$(iH1_YY) & Format(.Row, "00")
        Del_YYMME = RTrim$(iH1_YY) & Format(.RowSel, "00")
    End With
    
'  ( ämîFÉÅÉbÉZÅ[ÉW )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = " éÊà¯êÊ ÅÅ " & Trim(iH1_Torcd)
        
        gStr = " éÛì¸ ÅÅ " & RTrim$(Del_UkeireS)
        If Del_UkeireS <> Del_UkeireE Then gStr = gStr & " Å` " & RTrim$(Del_UkeireE)
        .MB_MSG(3) = gStr
        
        gStr = " îNåé ÅÅ " & RTrim$(Del_YYMMS)
        If Del_YYMMS <> Del_YYMME Then gStr = gStr & " Å` " & RTrim$(Del_YYMME)
        .MB_MSG(4) = gStr
        
        .MB_MSG(1) = "  Åy çÌèúèàóùämîF Åz      "
        
        .MB_Title = "çÌèúèàóù"
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Bd_SyoriKB = ""
            vsGrid1.ForeColor = vbBlack
            vsGrid1.BackColorSel = &H8000000D
            vsGrid1.RowSel = vsGrid1.Row
            vsGrid1.ColSel = vsGrid1.Col
            Exit Function
        End If
    End With
    
' ( çÌèúèàóù:ÉJÉåÉìÉ_É}ÉXÉ^ )
    gSL_Select = "Delete ÉJÉåÉìÉ_É}ÉXÉ^"
    gSL_Select = gSL_Select & " WHERE éÊà¯êÊ = " & "'" & RTrim$(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   éÊà¯êÊãÊï™ = " & CisFun.Val2(H1Op_Torkb(0).Tag)
    gSL_Select = gSL_Select & " And   îN = " & "'" & RTrim$(iH1_YY) & "'"
    
    CisDB.DBTran TransBegin                   ' ƒ◊›ªﬁ∏ºÆ›äJén
    On Error GoTo Uke_DBDelete_Err
    
    With vsGrid1
        For I = .Row To .RowSel
            For J = .Col To .ColSel
            
                SL_Select = gSL_Select
                SL_Select = SL_Select & " And   åé   = '" & Format(I, "00") & "'"
                SL_Select = SL_Select & " And   éÛì¸ = '" & RTrim$(.TextMatrix(0, J)) & "'"
            
                With CisDB
                    .SQL = SL_Select
                    .DBExec
                    
                    cnt_DEL = cnt_DEL + .RecordCount
                End With
            Next J
        Next I
    End With
    
    On Error GoTo 0
    CisDB.DBTran TransCommit                   ' ƒ◊›ªﬁ∏ºÆ›äJén
    
    iT1_Delsu = Format(cnt_DEL, "#,###")
    Uke_DBDelete = True
    Bd_SyoriKB = ""
    vsGrid1.ForeColor = vbBlack
    vsGrid1.BackColorSel = &H8000000D
    Exit Function

Uke_DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'*------------------------------------*
'*      éÊà¯êÊÅ@ó™èÃÅEçHèÍñºéQè∆        *
'*------------------------------------*
Private Function GetTorcd()
    GetTorcd = False
    
    If RTrim$(iH1_Torcd) = "" Then
        H1lb_Tornm = ""
        H1lb_Kojyonm = "ÅôÅ@é©é–ÉJÉåÉìÉ_Å@Åô"
        GoTo GetTorcd_ED
    End If
    
    SL_Select = "Select ó™èÃ, çHèÍñº From éÊà¯êÊÉ}ÉXÉ^"
    SL_Select = SL_Select & " Where éÊà¯êÊCD = '" & RTrim$(iH1_Torcd) & "'"
    SL_Select = SL_Select & " And   éÊà¯êÊãÊï™ = " & CisFun.Val2(H1Op_Torkb(0).Tag)
    
    If Not TRMRead(SL_Select, 1) Then
        H1lb_Tornm = ""
        H1lb_Kojyonm = ""
    Else
        H1lb_Tornm = RTrim$(TRM.ó™èÃ)
        H1lb_Kojyonm = RTrim$(TRM.çHèÍñº)
    End If

GetTorcd_ED:
    GetTorcd = True
End Function
'ÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅû
'ÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅû
'Åû
'Åû                         â“Å@ìÆÅ@ì˙Å@êîÅ@ÉOÅ@ÉäÅ@ÉbÅ@ÉhÅ@ä÷Å@åWÅ@èàÅ@óù
'Åû
'ÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅû
'ÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅû
'*--------------------------------------------------------------*
'*
'* [ â“ì≠ì˙êîÉOÉäÉbÉhÅ@óÒêîämíË ]
'*
'*          éÊà¯êÊéÛì¸É}ÉXÉ^ÇÊÇËÉOÉäÉbÉhÇÃóÒ(Cols)ämíË
'*
'*--------------------------------------------------------------*
Private Sub GridHead_Set(Optional pKB As String = "INV")
    Dim ID      As Integer
    
    w_Cols = c_Cols

    gSL_Select = "SELECT éÛì¸ FROM éÊà¯êÊéÛì¸É}ÉXÉ^"
'    gSL_Select = gSL_Select & " WHERE éÊà¯êÊCD = '" & RTrim$(iH1_Torcd) & "'"  ' 2006/03/28 Delete
'-------------------------------------------------- 2006/03/28 Insert
    If RTrim$(iH1_Torcd) = "" And H1Op_Torkb(1).Value Then
       gSL_Select = gSL_Select & " WHERE éÊà¯êÊCD = '" & RTrim$(gCompany) & RTrim(mJisyaKJ) & "'"
    Else
       gSL_Select = gSL_Select & " WHERE éÊà¯êÊCD = '" & RTrim$(iH1_Torcd) & "'"
    End If
'-------------------------------------------------- 2006/03/28 Insert End
    gSL_Select = gSL_Select & " AND   éÊà¯êÊãÊï™ = " & CisFun.Val2(H1Op_Torkb(0).Tag)
    gSL_Select = gSL_Select & " ORDER BY éÛì¸ "

    Ukeire_Su = 1
    
    If Not TUMRead(gSL_Select, , 1) Then
        ReDim CL_Ukeire(1)
        vsGrid1.Cols = w_Cols
        Call GridInit
    Else
        If CisDB.RecordCount > (w_Cols - 3) Then
            w_Cols = CisDB.RecordCount + 3
            vsGrid1.Cols = w_Cols
        End If
        
        Call GridInit

        ReDim CL_Ukeire(CisDB.RecordCount + 1)
        ID = 1
        
        vsGrid1.Redraw = False
        Do Until Not TUM_RDSTS
            With vsGrid1
                ID = ID + 1
                .TextMatrix(0, ID) = RTrim$(TUM.éÛì¸)
                CL_Ukeire(ID) = RTrim$(TUM.éÛì¸)
            End With
            Call TUMReadNext(1)
        Loop
        vsGrid1.Redraw = True
        
        Ukeire_Su = ID
    End If

    Call TUMClose(1)
End Sub
'*-------------------------------------------------------*
'*      â“Å@ì≠Å@ì˙Å@êîÅ@ÉOÅ@ÉäÅ@ÉbÅ@ÉhÅ@Å@ÉZÅ@ÉbÅ@Ég        *
'*-------------------------------------------------------*
Private Sub GridSet()
    Dim ID          As Integer
 
' íäèoèàóù èâä˙æØƒ
    Call CisFrm.MousePT(0)                                      ' œ≥ΩŒﬂ≤›¿ (çªéûåv)
    PNL_PF.Enabled = False                                      ' Ãß›∏ºÆ›égópã÷é~

    vsGrid1.Redraw = False
    Do Until Not CLM_RDSTS
        With CLM
            If .éÊà¯êÊãÊï™ <> CisFun.Val2(H1Op_Torkb(0).Tag) Or _
               RTrim$(.éÊà¯êÊ) <> RTrim$(iH1_Torcd) Or _
               .îN <> RTrim$(iH1_YY) Then Exit Do
        End With
        
        With vsGrid1
            
            For I = 1 To UBound(CL_Ukeire)
                If RTrim$(CL_Ukeire(I)) = RTrim$(CLM.éÛì¸) Then
                    .TextMatrix(CisFun.Val2(CLM.åé), I) = CLM.â“ìÆì˙êî
                    Exit For
                End If
            Next I
            
        End With
        Call CLMReadNext
    Loop
    
    vsGrid1.Redraw = True
        
GridSet_ED:
    PNL_PF.Enabled = True                                       ' Ãß›∏ºÆ›égópã÷é~ âèú
    Call CisFrm.MousePT(0)                                      ' œ≥ΩŒﬂ≤›¿ (çªéûåv)âèú
End Sub
'*-------------------------------------------------------*
'*      â“Å@ì≠Å@ì˙Å@êîÅ@ÉOÅ@ÉäÅ@ÉbÅ@ÉhÅ@Å@çƒÅ@ï\Å@é¶        *
'*-------------------------------------------------------*
Private Sub GridReDisp()
    
    CisVsGrid.Clear
    Call GridInit
    Call GridHead_Set
    
    If Not CLMRead(HD_Select) Then Exit Sub
    
    Call GridSet
    Call CLMClose
End Sub
'+----------------------------------------------------------+
'+      ÉOÉäÉbÉhèàóù                                         +
'+----------------------------------------------------------+
'******************************************'
'*****    å©èoÇµæØƒ & èâä˙èÛë‘æØƒ      *****
'******************************************'
Private Sub GridInit()
    Static sw_pStr          As String
    
    With CisVsGrid
        .Clear

        .SelectCol = w_Cols - 1
        .Init
        
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç»Çµ, 4, 0, "No")
        sw_pStr = "0"
        For I = 1 To w_Cols - 2
            Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 4, I, "")
            sw_pStr = sw_pStr & "," & I
        Next I
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 4, w_Cols - 1, "Select")

        .PatanSet 0, sw_pStr

        .InitDisp
    End With
  
End Sub
'*--------------------------------------------------------*
'*      vsGrid ( â“ìÆì˙êîÉOÉäÉbÉhÅ@É_ÉuÉãÉNÉäÉbÉN )         *
'*--------------------------------------------------------*
Private Sub VSGrid1_DblClick()
    SendKeys "{ENTER}"
End Sub
'ÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõ
'ÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõ
'Åõ
'Åõ                                 ÉJÅ@ÉåÅ@ÉìÅ@É_Å@ë§Å@ä÷Å@åWÅ@èàÅ@óù
'Åõ
'ÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõ
'ÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõÅúÅõ
'*-------------------------------------------------------------------*
'*                      É{Å@ÉfÅ@ÉBÅ@ÇQÅ@ï\Å@é¶
'*-------------------------------------------------------------------*
Private Sub Body2Set(Optional pCol As Integer = 0)

    If pCol = 0 Then pCol = vsGrid1.Col
    
    PB_Hiru(Save_Idx).BackColor = c_LFGray
    PB_Yakin(Save_Idx).BackColor = c_LFGray
    
    Static sw_Str           As String
'< ï\é¶ëŒè€îNåé >
    With CisFun
        .DateE_BefVal = RTrim$(iH1_YY) & Format(vsGrid1.Row, "00") & "01"
        .DateEdit
        gDate = .DateE_AfVal
    
        .DateEdit J_YM
        sw_Str = .DateE_AfVal
    End With
        
'< ÉJÉåÉìÉ_ïî ( ì˙Ç…Çø ) ÉZÉbÉg >
    If RTrim$(LB_Week(0).Tag) <> Format(gDate, "YYYYMM") Then Call CalendarDisplay

'< îNåéÉ^ÉCÉgÉãïî >
    With vsGrid1
        If RTrim$(.TextMatrix(0, pCol)) <> "" Then
            sw_Str = sw_Str & " éÛì¸ÅF" & RTrim$(.TextMatrix(0, pCol))
        End If
    End With
    lb_YYMM = sw_Str
    lb_YYMM.Tag = pCol

'< ÉJÉåÉìÉ_É}ÉXÉ^ >
    gSL_Select = "Select â“ìÆì˙êî, íããŒì˙êî, ñÈãŒì˙êî, â“ìÆãÊï™, ñÈãŒãÊï™"
    gSL_Select = gSL_Select & ", ó’éûíããÊï™, ó’éûñÈãÊï™"                                            '2005.02.28 Add
    gSL_Select = gSL_Select & " From ÉJÉåÉìÉ_É}ÉXÉ^"
    gSL_Select = gSL_Select & " Where éÊà¯êÊãÊï™ =  " & CisFun.Val2(H1Op_Torkb(0).Tag)
    gSL_Select = gSL_Select & " And   éÊà¯êÊ     = '" & RTrim$(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   îN         = '" & Format(gDate, "YYYY") & "'"
    gSL_Select = gSL_Select & " And   åé         = '" & Format(gDate, "MM") & "'"
    gSL_Select = gSL_Select & " And   éÛì¸       = '" & RTrim$(vsGrid1.TextMatrix(0, pCol)) & "'"
    gSL_Select = gSL_Select & " Order By éÊà¯êÊ, îN, åé, éÛì¸"
    
    If Not CLMRead(gSL_Select, 1) Then
'       < ÉfÅ[É^Ç»ÇµéûÅFìyì˙ãxÇ› >
        With CisFun
            .DateEdit
            gDate = .DateE_AfVal
        
'           ( ññì˙ )
            .Date_RsFormat = fDD
            .Date_End .DateE_BefVal
            gCnt2 = .Val2(.Date_Result)
        End With
    
        gStr = "": gCnt3 = 0
        gStr1 = ""                          '2005.02.28 Add
        For gCnt1 = 1 To gCnt2
            Select Case Format(gDate, "w")
                Case 1, 7:      gStr = gStr & "1"
                Case Else:      gStr = gStr & "0":  gCnt3 = gCnt3 + 1
            End Select
            gStr1 = gStr1 & "0"             '2005.02.28 Add
        
            gDate = DateAdd("d", 1, gDate)
        Next gCnt1
    
        CLM.â“ìÆì˙êî = gCnt3
        CLM.íããŒì˙êî = gCnt3
        CLM.ñÈãŒì˙êî = gCnt3
        CLM.â“ìÆãÊï™ = gStr
        CLM.ñÈãŒãÊï™ = gStr
        CLM.ó’éûíããÊï™ = gStr1              '2005.02.28 Add
        CLM.ó’éûñÈãÊï™ = gStr1              '2005.02.28 Add
    End If

    Static sw_STDD              As Byte

    sw_STDD = 99

''2005.02.28 Upd Start
'''< ÅuãxÅvÉZÉbÉg  >
''    For gCnt1 = 0 To 36
''
''        If B1lb_DD(gCnt1).Visible And sw_STDD = 99 Then sw_STDD = gCnt1
''
''        gCnt2 = CisFun.Val2(B1lb_DD(gCnt1))
''
''        PB_Hiru(gCnt1) = "":    PB_Yakin(gCnt1) = ""
''        If gCnt2 <> 0 Then
''            If Mid$(CLM.â“ìÆãÊï™, gCnt2, 1) = "1" Then PB_Hiru(gCnt1) = "ãx"
''            If Mid$(CLM.ñÈãŒãÊï™, gCnt2, 1) = "1" Then PB_Yakin(gCnt1) = "ãx"
''        End If
''    Next gCnt1
''
''    B1lb_Nisu(0) = CLM.íããŒì˙êî
''    B1lb_Nisu(1) = CLM.ñÈãŒì˙êî
''    B1lb_Nisu(2) = Format(CLM.â“ìÆì˙êî, "#0.0")
''
''    Save_Idx = sw_STDD
''    Save_HYkb = 0
''    PB_Hiru(Save_Idx).BackColor = c_FYellow
'< ÅuãxÅvÅuó’ÅvÉZÉbÉg  >
    For gCnt1 = 0 To 36
    
        If B1lb_DD(gCnt1).Visible And sw_STDD = 99 Then sw_STDD = gCnt1
        
        gCnt2 = CisFun.Val2(B1lb_DD(gCnt1))
        
        PB_Hiru(gCnt1) = "":    PB_Yakin(gCnt1) = ""
        If gCnt2 <> 0 Then
            If Mid$(CLM.â“ìÆãÊï™, gCnt2, 1) = "1" Then
                If Mid$(CLM.ó’éûíããÊï™, gCnt2, 1) <> "2" Then
                    PB_Hiru(gCnt1) = "ãx"
                    PB_Hiru(gCnt1).ForeColor = c_CRed               '2005.03.10 Add
                Else
                    PB_Hiru(gCnt1) = "ó’"
                    PB_Hiru(gCnt1).ForeColor = c_CGreen             '2005.03.10 Add
                End If
            End If
            If Mid$(CLM.ñÈãŒãÊï™, gCnt2, 1) = "1" Then
                If Mid$(CLM.ó’éûñÈãÊï™, gCnt2, 1) <> "2" Then
                    PB_Yakin(gCnt1) = "ãx"
                    PB_Yakin(gCnt1).ForeColor = c_CRed              '2005.03.10 Add
                Else
                    PB_Yakin(gCnt1) = "ó’"
                    PB_Yakin(gCnt1).ForeColor = c_CGreen            '2005.03.10 Add
                End If
            End If
        End If
    Next gCnt1

    B1lb_Nisu(0) = CLM.íããŒì˙êî
    B1lb_Nisu(1) = CLM.ñÈãŒì˙êî
    B1lb_Nisu(2) = Format(CLM.â“ìÆì˙êî, "#0.0")

    Save_Idx = sw_STDD
    Save_HYkb = 0
    PB_Hiru(Save_Idx).BackColor = c_FYellow
''2005.02.28 Upd End

'< É{ÉfÉBïîÇÃèàóùãÊï™(Bd_SyoriKB)ÉZÉbÉg >
    With vsGrid1
        gStr = ""
        For gInt = .Col To .ColSel
            If RTrim$(.TextMatrix(.Row, gInt)) = "" Then
                gStr = gStr & "A"
            Else
                gStr = gStr & "M"
            End If
        Next gInt

        If RTrim$(Replace(gStr, "A", "")) = "" Then Bd_SyoriKB = "ADD": Exit Sub
        If RTrim$(Replace(gStr, "M", "")) = "" Then Bd_SyoriKB = "MNT": Exit Sub
        Bd_SyoriKB = ""
    End With
End Sub
'*-------------------------------------------------------------------*
'*          ÉJÅ@ÉåÅ@ÉìÅ@É_Å@çÏÅ@ê¨Å@(Å@äOäœÅFì˙Ç…ÇøÅEà íuÅ@)            *
'*-------------------------------------------------------------------*
Private Function CalendarDisplay()
    Static sw_StWeek        As Integer
    Static sw_VisFlg        As Boolean
    Static sw_DDST          As Integer
    Static sw_DDED          As Integer

'< ï\é¶îNåéÇÃÇPì˙ÇÃójì˙êîílÇéZèo >
    sw_StWeek = Format(gDate, "w")

'< ï\é¶äJénà íu(ìYÇ¶éö) [ sw_DDST ] >   ÇPì˙
    sw_DDST = sw_StWeek - st_Week
    If sw_DDST < 0 Then sw_DDST = 7 + (sw_StWeek - st_Week)
    
'< ï\é¶èIóπà íu(ìYÇ¶éö) [ sw_DDED ] >   ññì˙
    gDate = DateAdd("m", 1, gDate)
    gDate = DateAdd("d", -1, gDate)
    sw_DDED = CisFun.Val2(Format(gDate, "D"))
    sw_DDED = sw_DDED + sw_DDST - 1
    
    gCnt1 = 0: gCnt2 = sw_StWeek - 1
    For gInt = 0 To 36
        sw_VisFlg = (sw_DDST <= gInt And gInt <= sw_DDED)
        
        B1lb_DD(gInt) = ""
        B1lb_DD(gInt).Visible = sw_VisFlg
        PB_Hiru(gInt).Visible = sw_VisFlg
        PB_Yakin(gInt).Visible = sw_VisFlg
        
        If sw_VisFlg Then
            gCnt1 = gCnt1 + 1
            B1lb_DD(gInt) = gCnt1
        
            gCnt2 = gCnt2 + 1
            If gCnt2 > 7 Then gCnt2 = 1
        
            Select Case gCnt2
                Case 1:         B1lb_DD(gInt).ForeColor = vbRed
                Case 7:         B1lb_DD(gInt).ForeColor = c_SatDBlue
                Case Else:      B1lb_DD(gInt).ForeColor = vbBlack
            End Select
        End If
    Next gInt

    LB_Week(0).Tag = Format(gDate, "YYYYMM")
End Function
'Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†
'Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†
'Å†
'Å†                 ÉJ Éå Éì É_ [âE]ë§ (BODY2) êß å‰ íÜ ÇÃ â“ìÆì˙êîÉOÉäÉbÉh [ç∂]ë§ ëÄ çÏ èà óù
'Å†
'Å†                     [ NumKey Ç…ÇÊÇÈëÄçÏ ]
'Å†                     [2]: éüåéÅA[8]:ëOåéÅA[4]:ç∂(éÛì¸à⁄ìÆ)ÅA[6]:âE(éÛì¸à⁄ìÆ)
'Å†
'Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†
'Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†
Private Sub DataChange(Key_Code As Integer)
    Static sw_OneSell       As Boolean
    Static sw_Col           As Integer
    
    Call CisFrm.MousePT(11)
    
    With vsGrid1
        sw_OneSell = (.Col = .ColSel)
        sw_Col = 0
        
        Select Case Key_Code
            Case vbKeyNumpad2
                                If .Row = 12 Then .Row = 1 Else .Row = .Row + 1
            Case vbKeyNumpad8
                                If .Row = 1 Then .Row = 12 Else .Row = .Row - 1
            Case vbKeyNumpad4
                                If sw_OneSell Then
                                    If .Col = 1 Then .Col = Ukeire_Su Else .Col = .Col - 1
                                Else
                                    sw_Col = lb_YYMM.Tag
                                    If sw_Col = 1 Then sw_Col = Ukeire_Su Else sw_Col = sw_Col - 1
                                End If
            Case vbKeyNumpad6
                                If sw_OneSell Then
                                    If .Col = Ukeire_Su Then .Col = 1 Else .Col = .Col + 1
                                Else
                                    sw_Col = lb_YYMM.Tag
                                    If sw_Col = Ukeire_Su Then sw_Col = 1 Else sw_Col = sw_Col + 1
                                End If
        End Select
    
        If Not .ColIsVisible(.Col) Then .LeftCol = .Col
    End With
    Call Body2Set(sw_Col)
    
    Call CisFrm.MousePT(1)
End Sub
'Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†
'Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†
'Å†
'Å†                     ñÓ àÛ ÉL Å[ Ç… ÇÊ ÇÈ ÉJ Éå Éì É_ [âE]ë§ (BODY2) ì‡ ÇÃ ì˙ Ç… Çø à⁄ ìÆ
'Å†
'Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†
'Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†
'+-----------------------------+
'+    ÉtÉHÅ[ÉJÉXà⁄ìÆ Å®Å©Å™Å´   +
'+-----------------------------+
Private Sub Idou_Forcus(Idou As Integer)
    
    PB_Hiru(Save_Idx).BackColor = c_LFGray
    PB_Yakin(Save_Idx).BackColor = c_LFGray
    
    If Idou = vbKeyLeft And Save_HYkb = 1 Then Save_HYkb = 0: GoTo Idou_Forcus_ED
    If Idou = vbKeyRight And Save_HYkb = 0 Then Save_HYkb = 1: GoTo Idou_Forcus_ED
    
    Static sw_Zogen         As Integer
    Static sw_Val           As Integer
    Static sw_End           As Integer
    Static sw_This          As Integer
    
    Select Case Idou
        Case vbKeyUp:       sw_Zogen = -7:  sw_End = 36
        Case vbKeyDown:     sw_Zogen = 7:   sw_End = 0
        Case vbKeyLeft:     sw_Zogen = -1:  sw_End = 36:    Save_HYkb = 1
        Case vbKeyRight:    sw_Zogen = 1:   sw_End = 0:     Save_HYkb = 0
    End Select

    sw_Val = Save_Idx + sw_Zogen

'   < ëOå„ç∂âE >
    If sw_Val >= 0 And sw_Val <= 36 Then
        If B1lb_DD(sw_Val).Visible Then
            Save_Idx = sw_Val
            GoTo Idou_Forcus_ED
        End If
    End If
    
'   < ëOå„ç∂âEÇ≈ÇÕÇ»Ç≠ÅAç≈èâÅEç≈èIÇíTÇ∑ >
    sw_This = 99
    sw_Zogen = sw_Zogen * -1
    For I = Save_Idx + sw_Zogen To sw_End Step sw_Zogen
        If B1lb_DD(I).Visible Then sw_This = I
    Next I
    
    If sw_This = 99 Then Exit Sub
    Save_Idx = sw_This
    
Idou_Forcus_ED:
    If Save_HYkb = 0 Then
        PB_Hiru(Save_Idx).BackColor = c_FYellow
    Else
        PB_Yakin(Save_Idx).BackColor = c_FYellow
    End If
End Sub
'Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§
'Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§
'Å§
'Å§                         ñºÅ@èÃÅ@É}Å@ÉXÅ@É^Å@ämÅ@îF  ( ÉV ÉX Ée ÉÄ ã§ í  íË êî ìô )
'Å§
'Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§
'Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§Å£Å§
'*--------------------------------------------------------------------*
'*      ÉJ  Éå  Éì  É_  ój  ì˙  óì  èÓ  ïÒ  äl  ìæ  Åï  ÉZ  Éb  Ég      *
'*--------------------------------------------------------------------*
Private Function DayOfTheWeek_Set()
    
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
    
    Call SYMInitCreate(False)

    Call ItemsClearSYM
    gSL_Select = "Select êîéö1 From ñºèÃÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " Where ãÊï™ñºèÃ = 'ójì˙äJénä«óù'"
    gSL_Select = gSL_Select & " And   íl = '1'"
    Call SYMRead(gSL_Select, 1)
    
    st_Week = SYM.êîéö1 + 1

    For I = 0 To 6
        gInt = I + st_Week
        If gInt > 7 Then gInt = gInt - 7
    
'       < ójì˙ >
        LB_Week(I) = Format(gInt, "aaa")
        
'       < êFê›íË >
        Select Case gInt
            Case 1:         LB_Week(I).BackColor = vbRed
            Case 7:         LB_Week(I).BackColor = c_SatDBlue
            Case Else:      LB_Week(I).BackColor = c_WeekEndD
        End Select
    Next I
End Function

