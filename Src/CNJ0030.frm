VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{8F4FF474-3E18-11D4-AEC2-00000EA57F9E}#1.0#0"; "CISMSG~1.OCX"
Begin VB.Form CNJ0030 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "î≠íçì‡é¶ì¸óÕ"
   ClientHeight    =   11235
   ClientLeft      =   2040
   ClientTop       =   1755
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
      TabIndex        =   72
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
               TabIndex        =   73
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
            Left            =   10680
            Top             =   825
            Width           =   1545
            _ExtentX        =   2725
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
               Left            =   75
               Top             =   360
               Width           =   1380
               _ExtentX        =   2434
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
            Left            =   8670
            Top             =   30
            Width           =   840
            _ExtentX        =   1482
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
               Left            =   90
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
            Left            =   7830
            Top             =   30
            Width           =   840
            _ExtentX        =   1482
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
               Left            =   90
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
            Left            =   12225
            Top             =   825
            Width           =   1545
            _ExtentX        =   2725
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
               Left            =   75
               Top             =   360
               Width           =   1380
               _ExtentX        =   2434
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
            Left            =   5865
            Top             =   825
            Width           =   3270
            _ExtentX        =   5768
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
               Left            =   75
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
               Left            =   1815
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
               Left            =   1485
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
            Left            =   9510
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
               Left            =   120
               Top             =   390
               Width           =   1425
               _ExtentX        =   2514
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
            Left            =   6990
            Top             =   30
            Width           =   840
            _ExtentX        =   1482
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
               Left            =   90
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
            Left            =   5040
            Top             =   825
            Width           =   825
            _ExtentX        =   1455
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
            Left            =   9135
            Top             =   825
            Width           =   1545
            _ExtentX        =   2725
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
               Left            =   75
               Top             =   360
               Width           =   1380
               _ExtentX        =   2434
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
         Begin Cis3D_v60.CIS3D CIS3D29 
            Height          =   795
            Left            =   1380
            Top             =   30
            Width           =   1260
            _ExtentX        =   2223
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
               Left            =   120
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
            Left            =   2640
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
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
               Left            =   90
               Top             =   390
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   582
               BackColor       =   12648384
               Caption         =   "99,999"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Left            =   3870
            Top             =   30
            Width           =   1890
            _ExtentX        =   3334
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
               Left            =   90
               Top             =   390
               Width           =   1275
               _ExtentX        =   2249
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
               Left            =   1380
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
            Left            =   11190
            Top             =   30
            Width           =   2580
            _ExtentX        =   4551
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
               Left            =   120
               Top             =   390
               Width           =   1155
               _ExtentX        =   2037
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
               Left            =   1320
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
            Left            =   5760
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
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
               Left            =   90
               Top             =   390
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   582
               BackColor       =   12648384
               Caption         =   "99,999"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               TabIndex        =   76
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
            Left            =   930
            Top             =   825
            Width           =   4110
            _ExtentX        =   7250
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
               Height          =   690
               Left            =   2040
               TabIndex        =   74
               TabStop         =   0   'False
               Top             =   60
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   1217
               BackColor       =   -2147483633
               ForeColor       =   12582912
               Caption         =   "999,999,999"
               BackStyle       =   0
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
               Height          =   690
               Index           =   0
               Left            =   495
               TabIndex        =   75
               TabStop         =   0   'False
               Top             =   60
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   1217
               BackColor       =   -2147483633
               ForeColor       =   12582912
               Caption         =   "åªî≠íççœêî"
               BackStyle       =   0
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
               gGM1_Red        =   3
               gGM2_Green      =   3
               gGradation      =   -1  'True
            End
         End
      End
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   1560
      Left            =   720
      ScaleHeight     =   1500
      ScaleWidth      =   9045
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   900
      Width           =   9105
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   495
         Left            =   2820
         Top             =   75
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   873
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
         cAlingnment     =   8
         cPositionY      =   -30
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   855
         Left            =   60
         Top             =   570
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   1508
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
      Begin Cis3D_v60.CIS3D BLB_HINBN 
         Height          =   855
         Left            =   1500
         Top             =   570
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   1508
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
         cPositionY      =   45
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   390
            Width           =   3210
            _ExtentX        =   5662
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
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   855
         Left            =   4890
         Top             =   570
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   1508
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
            TabIndex        =   3
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
            Left            =   3555
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
            Left            =   3630
            TabIndex        =   4
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
         Height          =   495
         Left            =   60
         Top             =   75
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   873
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
         cAlingnment     =   1
         cPositionX      =   240
         Begin CisYM_V60.CisYM iH1_NJYM 
            Height          =   375
            Left            =   1350
            TabIndex        =   5
            Top             =   60
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
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   1305
      Left            =   10373
      ScaleHeight     =   1245
      ScaleWidth      =   4215
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1155
      Width           =   4275
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1695
         Top             =   600
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
         Left            =   2520
         Top             =   600
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
         Left            =   3345
         Top             =   600
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
         Left            =   45
         Top             =   600
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
         Left            =   870
         Top             =   600
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
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "Åy î≠íçì‡é¶ì¸óÕ  Åz"
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
         TabIndex        =   41
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
         Left            =   14400
         Top             =   60
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
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   285
         Left            =   8100
         Top             =   30
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
      TabIndex        =   37
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
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   4515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   31
                  Left            =   -15
                  TabIndex        =   36
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   70
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   30
                  Left            =   -15
                  TabIndex        =   35
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   29
                  Left            =   -15
                  TabIndex        =   34
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   68
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   28
                  Left            =   -15
                  TabIndex        =   33
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   67
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   27
                  Left            =   -15
                  TabIndex        =   32
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   26
                  Left            =   -15
                  TabIndex        =   31
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   25
                  Left            =   -15
                  TabIndex        =   30
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   24
                  Left            =   -15
                  TabIndex        =   29
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   23
                  Left            =   -15
                  TabIndex        =   28
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   22
                  Left            =   -15
                  TabIndex        =   27
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   21
                  Left            =   -15
                  TabIndex        =   26
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   20
                  Left            =   -15
                  TabIndex        =   25
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   19
                  Left            =   -15
                  TabIndex        =   24
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   18
                  Left            =   -15
                  TabIndex        =   23
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   17
                  Left            =   -15
                  TabIndex        =   22
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   16
                  Left            =   -15
                  TabIndex        =   21
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   15
                  Left            =   -15
                  TabIndex        =   20
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   14
                  Left            =   -15
                  TabIndex        =   19
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   13
                  Left            =   -15
                  TabIndex        =   18
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   12
                  Left            =   -15
                  TabIndex        =   17
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   11
                  Left            =   -15
                  TabIndex        =   16
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   10
                  Left            =   -15
                  TabIndex        =   15
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   9
                  Left            =   -15
                  TabIndex        =   14
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   8
                  Left            =   -15
                  TabIndex        =   13
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   7
                  Left            =   -15
                  TabIndex        =   12
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   6
                  Left            =   -15
                  TabIndex        =   11
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   5
                  Left            =   -15
                  TabIndex        =   10
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   4
                  Left            =   -15
                  TabIndex        =   9
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   3
                  Left            =   -15
                  TabIndex        =   8
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   2
                  Left            =   -15
                  TabIndex        =   7
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Left            =   1740
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
                  Left            =   750
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
                  Left            =   2685
                  Top             =   210
                  Width           =   3045
                  _ExtentX        =   5371
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
                     Left            =   1110
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
                     Left            =   1470
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
                     Left            =   1830
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
                     Left            =   2190
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
                     Left            =   2550
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
                  Left            =   780
                  Top             =   930
                  Width           =   1845
                  _ExtentX        =   3254
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
                     Left            =   960
                     Top             =   15
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   582
                     BackColor       =   12648384
                     Caption         =   "99,999"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   3180
                  Top             =   930
                  Width           =   2535
                  _ExtentX        =   4471
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
                     Left            =   600
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
                     Left            =   960
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
                     Left            =   1320
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
                     Left            =   1680
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
                     Left            =   2040
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
                  Width           =   2415
               End
               Begin VB.Shape Shape3 
                  Height          =   675
                  Left            =   2940
                  Shape           =   4  'ä€Ç›ÇÃÇ†ÇÈí∑ï˚å`
                  Top             =   750
                  Width           =   2865
               End
               Begin VB.Shape Shape2 
                  Height          =   675
                  Left            =   2430
                  Shape           =   4  'ä€Ç›ÇÃÇ†ÇÈí∑ï˚å`
                  Top             =   60
                  Width           =   3375
               End
               Begin VB.Shape Shape1 
                  Height          =   675
                  Left            =   510
                  Shape           =   4  'ä€Ç›ÇÃÇ†ÇÈí∑ï˚å`
                  Top             =   60
                  Width           =   1905
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
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   15
               Width           =   1245
               Begin CisText_V60.CisText iB1_Su 
                  Height          =   315
                  Index           =   1
                  Left            =   -15
                  TabIndex        =   6
                  Top             =   855
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  cFormat         =   "##,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
                  Caption         =   "##,###,###"
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
               Caption         =   "99,999,999"
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
                  Caption         =   "#,###,###,###"
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
                  Caption         =   "#,###,###,###"
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
                  Caption         =   "#,###,###,###"
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
                  Caption         =   "#,###,###,###"
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
               Caption         =   "99,999,999"
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
               Caption         =   "99,999,999"
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
               Caption         =   "99,999,999"
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
                  Caption         =   "#,###,###,###"
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
                  Caption         =   "#,###,###,###"
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
               Caption         =   "99,999,999"
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
               Caption         =   "99,999,999"
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
Attribute VB_Name = "CNJ0030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :   î≠íçì‡é¶ì¸óÕ
'**       ÉtÉHÅ[ÉÄID    :   CNJ0030
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :   2004/03/22  By CIS
'**       ïœ  çX  ì˙    :   2004/04/09  ëOâÒÉfÅ[É^çXêV(ïœçXï™Å®ëSïî)
'**       ïœ  çX  ì˙    :   2005/03/08  èIóπÇÃéûÅAèàóùëŒè€ÉfÅ[É^Ç™ñ≥Ç¢éûÇ≈Ç‡èàóùãÊï™ÉNÉäÉAÇé¿çs
'**       ïœ  çX  ì˙    :   2005/03/19  è„ãLÇÃèCê≥É~ÉXëŒâû
'**       ïœ  çX  ì˙    :   2005/03/22  ì¸óÕŒﬁ¿›Çâüâ∫éûÅAÅwé¿çséû¥◊∞'400'ÅxÇï\é¶Ç∑ÇÈéûÇ™Ç†ÇÈÅBÅiÇPïiî‘Ç≈ï°êîéËîzêÊÇ™óLÇÈéûÅj
'**       ïœ  çX  ì˙    :   2005/04/05  ëŒè€îNåéÇÃ¡™Ø∏í«â¡Å®ì‡é¶îNåéÇ…ëŒÇµÇƒÅAìñåéÅEóÇåéÅEóÇÅXåéï™ÇÃÇRÉñåéÇÃÇ›ÇnÇjÇ∆Ç∑ÇÈÅB
'**
'   200312,200312,000001111123,0001A,01
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' Ã´∞—±∏√®Ãﬁ Ã◊∏ﬁ
    Dim ProcHB                  As String * 2   ' ì¸óÕàÊêÿë÷
    Dim SyoriKB                 As String       ' èàóùãÊï™(êVãK•èCê≥•çÌèú ìô)
    Dim SyoriNM                 As String       ' èàóùñºèÃ(êVãK•èCê≥•çÌèú ìô)

    Dim mOndayColor             As Long         ' â“ì≠ì˙
    Dim mOffdayColor            As Long         ' ãxì˙
    Dim mNInitColor             As Long         ' ì‡é¶
    Dim mNHenColor              As Long         ' ì‡é¶(ïœçX)
    Dim mHHenColor              As Long         ' î≠íçì‡é¶(ïœçX)

    Dim wShift                  As Integer      ' PageUp/PageDownéû ( ëŒè€îNåé )
    Dim sv_Choku                As String       ' íºëóêÊ(ëﬁîóp)
    Dim sv_Hinbn                As String       ' ïiî‘(ëﬁî)

    Dim wInt                    As Integer

'( ÉvÉçÉOÉâÉÄòAågóp )
    Dim cmd_KB                  As Byte
    Dim cmd_NJYM                As String
    Dim cmd_YM                  As String
    Dim cmd_Hinbn               As String
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

        If UBound(wSplit) <> 4 Then
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
        cmd_Hinbn = RTrim$(wSplit(2))
        cmd_Torcd = RTrim$(wSplit(3))
        cmd_Ukeir = RTrim$(wSplit(4))

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
    sv_Hinbn = ""

'< î≠íçì‡é¶îNåé >
    Call GetNJProc_YM(3)
    If RTrim$(gn_HNYM) = "" Then gn_HNYM = Format(Now(), "YYYYMM")
    iH1_NJYM = gn_HNYM
    iH1_YM = gn_HNYM

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
        iH1_Hinbn = cmd_Hinbn
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
'Å§Å§---------------------------------------------<< 2005/03/22 >> Update Start
'                                   Dummy.Enabled = True: Dummy.SetFocus: DoEvents
'                                   Call PB_ENT_Click
'                                   Dummy.Enabled = False
                                    Call PB_ENT_Click
'Å¢Å¢---------------------------------------------<< 2005/03/22 >> Update End
                                End If
           Case Else:           Exit Function
    End Select
    
    Key_Acc = True
End Function
'+---------------------+
'+    èI  óπ  ÉL Å[     +
'+---------------------+
Private Sub PB_END_Click()
'   < å„ï˚èàóù >
    If cmd_KB = 0 Then Call After_Proc
    
    Call CisDB.DBDISConnect                         ' √ﬁ∞¿Õﬁ∞Ωêÿíf
    Call CisFrm.UnLoadDisp                          ' Ã´∞—µÃﬁºﬁ™∏ƒè¡ãé
    End
End Sub
'+----------------------------+
'+    ì¸  óÕ  ÉL Å[(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
'Å§Å§---------------------------------------------<< 2005/03/22 >> Update Start
'   Call ReturnPress
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call ReturnPress
    Dummy.Enabled = False
'Å¢Å¢---------------------------------------------<< 2005/03/22 >> Update End
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
        gSL_Select = "Select ì‡é¶îNåé, åvâÊîNåé, ïiî‘, éËîzêÊ, éÛì¸ From î≠íçì‡é¶ÉeÅ[ÉuÉã"
        gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        gSL_Select = gSL_Select & " And   åvâÊîNåé = ì‡é¶îNåé"
        gSL_Select = gSL_Select & " And   SubString( ïiî‘   + Space(25), 1, 25 )"
        gSL_Select = gSL_Select & "     + SubString( éËîzêÊ + Space(7),  1,  7 )"
        gSL_Select = gSL_Select & "     + SubString( éÛì¸   + Space(2),  1,  2 )"
        gSL_Select = gSL_Select & "       < '"
        gSL_Select = gSL_Select & RTrim$(iH1_Hinbn) & Space(25 - Len(RTrim$(iH1_Hinbn)))
        gSL_Select = gSL_Select & RTrim$(iH1_Torcd) & Space(7 - Len(RTrim$(iH1_Torcd)))
        gSL_Select = gSL_Select & RTrim$(iH1_Ukeir) & Space(2 - Len(RTrim$(iH1_Ukeir))) & "'"
'        gSL_Select = gSL_Select & " ORDER BY SubString( ïiî‘   + Space(25), 1, 25 )"
'        gSL_Select = gSL_Select & "        + SubString( éËîzêÊ + Space(7),  1,  7 )"
'        gSL_Select = gSL_Select & "        + SubString( éÛì¸   + Space(2),  1,  2 ) Desc"
        gSL_Select = gSL_Select & " ORDER BY ïiî‘ Desc, éËîzêÊ Desc, éÛì¸ Desc"
    Else
        gSL_Select = "Select ì‡é¶îNåé, åvâÊîNåé, ïiî‘, éËîzêÊ, éÛì¸ From î≠íçì‡é¶ÉeÅ[ÉuÉã"
        gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim$(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim$(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim$(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & " And   åvâÊîNåé < '" & RTrim$(iH1_YM) & "'"
        gSL_Select = gSL_Select & " ORDER BY åvâÊîNåé Desc"
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
        gSL_Select = "Select ì‡é¶îNåé, åvâÊîNåé, ïiî‘, éËîzêÊ, éÛì¸ From î≠íçì‡é¶ÉeÅ[ÉuÉã"
        gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        gSL_Select = gSL_Select & " And   åvâÊîNåé = ì‡é¶îNåé"
        gSL_Select = gSL_Select & " And   SubString( ïiî‘   + Space(25), 1, 25 )"
        gSL_Select = gSL_Select & "     + SubString( éËîzêÊ + Space(7),  1,  7 )"
        gSL_Select = gSL_Select & "     + SubString( éÛì¸   + Space(2),  1,  2 )"
        gSL_Select = gSL_Select & "       > '"
        gSL_Select = gSL_Select & RTrim$(iH1_Hinbn) & Space(25 - Len(RTrim$(iH1_Hinbn)))
        gSL_Select = gSL_Select & RTrim$(iH1_Torcd) & Space(7 - Len(RTrim$(iH1_Torcd)))
        gSL_Select = gSL_Select & RTrim$(iH1_Ukeir) & Space(2 - Len(RTrim$(iH1_Ukeir))) & "'"
'        gSL_Select = gSL_Select & " ORDER BY SubString( ïiî‘   + Space(25), 1, 25 )"
'        gSL_Select = gSL_Select & "        + SubString( éËîzêÊ + Space(7),  1,  7 )"
'        gSL_Select = gSL_Select & "        + SubString( éÛì¸   + Space(2),  1,  2 )"
        gSL_Select = gSL_Select & " ORDER BY ïiî‘, éËîzêÊ, éÛì¸"
    Else
        gSL_Select = "Select ì‡é¶îNåé, åvâÊîNåé, ïiî‘, éËîzêÊ, éÛì¸ From î≠íçì‡é¶ÉeÅ[ÉuÉã"
        gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim$(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim$(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim$(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & " And   åvâÊîNåé > '" & RTrim$(iH1_YM) & "'"
        gSL_Select = gSL_Select & " ORDER BY åvâÊîNåé"
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
        iB1_Su(gInt) = CisFun.RSetFld(iB1_Su(gInt).Tag, 10, "#,###")
    Next gInt
    B1lb_GSu(0) = CisFun.RSetFld(B1lb_GSu(0).Tag, 11, "#,###")
    
    If wInt = 0 Then wInt = 1
    iB1_Su(wInt).SetFocus
    
    Dummy.Enabled = False
End Sub
'+----------------------------+
'+    åü  çı  ÉL Å[(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Tehacd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_Tehacd: Exit Sub
End Sub
'+-----------------------------------+
'+       ïiî‘åüçı                     +
'+-----------------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 1
    RVI_SK = 9
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iH1_Hinbn = RV_Hinbn
        iH1_Torcd = RV_TorcdK
        iH1_Ukeir = RV_Ukeir
    
        H1lb_Tornm = ""
        If TorNmGet(iH1_Torcd, 9) Then H1lb_Tornm = RTrim(TRM_RName)
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
    RV_Hinbn = RTrim(iH1_Hinbn)
    If SyoriKB = "ADD" Then RV_Kubun = "MST" Else RV_Kubun = "HNJ"
    
    CNJ0025.Show vbModal
    Unload CNJ0025
    Set CNJ0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_Torcd
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RV_Ukeir
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

    If RTrim$(iH1_Hinbn) <> sv_Hinbn And _
       ActiveControl.Name <> "Dummy" And _
       ActiveControl.Name <> "iH1_YM" Then
        Call SetTorcd
        sv_Hinbn = RTrim$(iH1_Hinbn)
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
    If TorNmGet(iH1_Torcd, 9) Then
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
    B1lb_GSu(0) = Format(gLong, "#,###")
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

' * édì¸êÊ é©ìÆï\é¶
    If RTrim$(iH1_Torcd) = "" Then Call SetTorcd

' * édì¸êÊ
    If Trim(iH1_Torcd) <> "" Then
        gSL_Select = " SELECT * "
        gSL_Select = gSL_Select & " FROM éÊà¯êÊÉ}ÉXÉ^ "
        gSL_Select = gSL_Select & " WHERE éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   And éÊà¯êÊãÊï™ = 1"
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
    
' * ïiî‘éÊà¯êÊÉ}ÉXÉ^
    gSL_Select = "Select ïiî‘ From ïiî‘éÊà¯êÊÉ}ÉXÉ^"
    gSL_Select = gSL_Select & " Where ïiî‘ = '" & RTrim$(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   éÊà¯êÊ = '" & RTrim$(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim$(iH1_Ukeir) & "'"
    If Not HTMRead(gSL_Select, 1) Then
         With CisFun
             .MB_Lines = 4
             .MB_MSG(2) = "    ïiî‘éÊà¯êÊÉ}ÉXÉ^ñ¢ìoò^            "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
         End With
        iH1_Hinbn.SetFocus
        Exit Function
    End If

    gSL_Select = "Select * From î≠íçì‡é¶ÉeÅ[ÉuÉã"
    gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    If DBInput("INV") Then
       If SyoriKB = "ADD" Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    î≠íçì‡é¶ÉfÅ[É^ ìoò^çœ            "
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
              .MB_MSG(2) = "    î≠íçì‡é¶ÉfÅ[É^ ñ¢ìoò^            "
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

' * ïiî‘
    If Trim(iH1_Hinbn) = "" Then GoTo Head1Chk_Cmd_Err

' * éËîzêÊ
    If Trim(iH1_Torcd) <> "" Then
        gSL_Select = " SELECT * "
        gSL_Select = gSL_Select & " FROM éÊà¯êÊÉ}ÉXÉ^ "
        gSL_Select = gSL_Select & " WHERE éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   And éÊà¯êÊãÊï™ = 1"
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
    gSL_Select = "Select * From î≠íçì‡é¶ÉeÅ[ÉuÉã"
    gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
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
    
    If Not HNJRead(gSL_Select, 1) Then
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
    Else
        iH1_NJYM = RTrim(HNJ.ì‡é¶îNåé)
        iH1_YM = RTrim(HNJ.åvâÊîNåé)
        iH1_Hinbn = RTrim(HNJ.ïiî‘)
        iH1_Torcd = RTrim(HNJ.éËîzêÊ)
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RTrim(HNJ.éÛì¸)

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
    gSL_Select = gSL_Select & " From î≠íçì‡é¶ïKóvêîëOâÒÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where ïiî‘     = '" & RTrim(iH1_Hinbn) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éËîzêÊ   = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸     = '" & RTrim(iH1_Ukeir) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'" & vbCrLf

    gSL_Select = gSL_Select & " Union ALL" & vbCrLf
    gSL_Select = gSL_Select & " Select '2' KB, SubString( î[ä˙, 7, 2 ) DD, êîó " & vbCrLf
    gSL_Select = gSL_Select & " From î≠íçì‡é¶ëOâÒÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where ïiî‘     = '" & RTrim(iH1_Hinbn) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éËîzêÊ   = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸     = '" & RTrim(iH1_Ukeir) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'" & vbCrLf

    gSL_Select = gSL_Select & " Union ALL" & vbCrLf
    gSL_Select = gSL_Select & " Select '3' KB, SubString( î[ä˙, 7, 2 ) DD, êîó " & vbCrLf
    gSL_Select = gSL_Select & " From î≠íçì‡é¶ÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where ïiî‘     = '" & RTrim(iH1_Hinbn) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éËîzêÊ   = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸     = '" & RTrim(iH1_Ukeir) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'" & vbCrLf

    gSL_Select = gSL_Select & " Union ALL" & vbCrLf
    gSL_Select = gSL_Select & " Select '4' KB, SubString( î[ä˙, 7, 2 ) DD, êîó " & vbCrLf
    gSL_Select = gSL_Select & " From î≠íçì‡é¶ïKóvêîÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where ïiî‘     = '" & RTrim(iH1_Hinbn) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éËîzêÊ   = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸     = '" & RTrim(iH1_Ukeir) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'" & vbCrLf
    
'   < åéó êî >
    gSL_Select = gSL_Select & " Union ALL" & vbCrLf
    gSL_Select = gSL_Select & " Select '1' KB,"
    gSL_Select = gSL_Select & " DateDiff( M, ì‡é¶îNåé + '01', åvâÊîNåé + '01' ) + 50 DD,"
    gSL_Select = gSL_Select & " Sum( êîó  ) êîó " & vbCrLf
    gSL_Select = gSL_Select & " From î≠íçì‡é¶ïKóvêîëOâÒÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where ïiî‘      = '" & RTrim(iH1_Hinbn) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éËîzêÊ    = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸      = '" & RTrim(iH1_Ukeir) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé >= '" & RTrim(iH1_YM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé  = '" & RTrim(iH1_NJYM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " Group By ì‡é¶îNåé, åvâÊîNåé" & vbCrLf
    
    gSL_Select = gSL_Select & " Union ALL" & vbCrLf
    gSL_Select = gSL_Select & " Select '2' KB,"
    gSL_Select = gSL_Select & " DateDiff( M, ì‡é¶îNåé + '01', åvâÊîNåé + '01' ) + 50 DD,"
    gSL_Select = gSL_Select & " Sum( êîó  ) êîó " & vbCrLf
    gSL_Select = gSL_Select & " From î≠íçì‡é¶ëOâÒÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where ïiî‘      = '" & RTrim(iH1_Hinbn) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éËîzêÊ    = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸      = '" & RTrim(iH1_Ukeir) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé >= '" & RTrim(iH1_YM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé  = '" & RTrim(iH1_NJYM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " Group By ì‡é¶îNåé, åvâÊîNåé" & vbCrLf

    gSL_Select = gSL_Select & " Union ALL" & vbCrLf
    gSL_Select = gSL_Select & " Select '3' KB,"
    gSL_Select = gSL_Select & " DateDiff( M, ì‡é¶îNåé + '01', åvâÊîNåé + '01' ) + 50 DD,"
    gSL_Select = gSL_Select & " Sum( êîó  ) êîó " & vbCrLf
    gSL_Select = gSL_Select & " From î≠íçì‡é¶ÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where ïiî‘      = '" & RTrim(iH1_Hinbn) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éËîzêÊ    = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸      = '" & RTrim(iH1_Ukeir) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé >= '" & RTrim(iH1_YM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé  = '" & RTrim(iH1_NJYM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " Group By ì‡é¶îNåé, åvâÊîNåé" & vbCrLf

    gSL_Select = gSL_Select & " Union ALL" & vbCrLf
    gSL_Select = gSL_Select & " Select '4' KB,"
    gSL_Select = gSL_Select & " DateDiff( M, ì‡é¶îNåé + '01', åvâÊîNåé + '01' ) + 50 DD,"
    gSL_Select = gSL_Select & " Sum( êîó  ) êîó " & vbCrLf
    gSL_Select = gSL_Select & " From î≠íçì‡é¶ïKóvêîÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where ïiî‘      = '" & RTrim(iH1_Hinbn) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éËîzêÊ    = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   éÛì¸      = '" & RTrim(iH1_Ukeir) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé >= '" & RTrim(iH1_YM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé  = '" & RTrim(iH1_NJYM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " Group By ì‡é¶îNåé, åvâÊîNåé" & vbCrLf
    
'   < ïiî‘íPà Å@ìñåéêîó  >
    gSL_Select = gSL_Select & " Union ALL" & vbCrLf
    gSL_Select = gSL_Select & " Select '3' KB, 70 DD, IsNull( Sum( êîó  ), 0 ) êîó " & vbCrLf
    gSL_Select = gSL_Select & " From î≠íçì‡é¶ÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where ïiî‘     = '" & RTrim(iH1_Hinbn) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'" & vbCrLf
    
'   < ïiî‘íPà Å@ëOåéï™ìñåéêîó  >
    gSL_Select = gSL_Select & " Union ALL" & vbCrLf
    gSL_Select = gSL_Select & " Select '3' KB, 71 DD, IsNull( Sum( êîó  ), 0 ) êîó " & vbCrLf
    gSL_Select = gSL_Select & " From î≠íçì‡é¶ÉeÅ[ÉuÉã" & vbCrLf
    gSL_Select = gSL_Select & " Where ïiî‘     = '" & RTrim(iH1_Hinbn) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   åvâÊîNåé = Convert( Varchar(6), DateAdd( M, -1, '" & RTrim$(iH1_YM) & "' + '01' ), 112 )" & vbCrLf
    gSL_Select = gSL_Select & " And   ì‡é¶îNåé = Convert( Varchar(6), DateAdd( M, -1, '" & RTrim$(iH1_NJYM) & "' + '01' ), 112 )" & vbCrLf
    
    gSL_Select = gSL_Select & " Order by KB, DD" & vbCrLf
    If HNJRead(gSL_Select) Then
        Do Until Not HNJ_RDSTS
            With HNJ

                If .DD <= 31 Then
'                   [ ì˙ó  ]
                    Select Case .KB
                        Case "1"
                                    B1lb_LNJSu(.DD) = Format(.êîó , "#,###")                '< ïKóvêî [ ëOâÒ ] >
                        Case "2"
                                    B1lb_LSu(.DD) = Format(.êîó , "#,###")                  '< åvâÊêî [ ëOâÒ ] >
                                    B1lb_LSu(.DD).Tag = "X"                                 '  çXêVéûÇ…Ç±ÇÃÉ^ÉOÇégóp
                        Case "3"
                                    iB1_Su(.DD) = CisFun.RSetFld(.êîó , 10, "#,###")        '< åvâÊêî >
                                    iB1_Su(.DD).Tag = .êîó 
                                    If .êîó  <> CisFun.Val2(B1lb_LSu(.DD)) And _
                                       B1lb_LSu(.DD).Tag = "X" Then B1lb_LSu(.DD).BackColor = mHHenColor
                        Case "4"
                                    B1lb_NJSu(.DD) = Format(.êîó , "#,###")                 '< ïKóvêî >
                    End Select
                Else
                    If .DD >= sw_MConst And .DD <= sw_MConst + 2 Then
'                       [ åéó  ]
                        Select Case .KB
                            Case "1":   B1lb_LGNJSu(.DD - sw_MConst) = Format(.êîó , "#,###")      '< ïKóvêî [ ëOâÒ ] >
                            Case "2":   B1lb_LGSu(.DD - sw_MConst) = Format(.êîó , "#,###")        '< åvâÊêî [ ëOâÒ ] >
                            Case "3"
                                        B1lb_GSu(.DD - sw_MConst) = Format(.êîó , "#,###")         '< åvâÊêî >
                                        B1lb_GSu(.DD - sw_MConst).Tag = .êîó 
                            Case "4":   B1lb_GNJSu(.DD - sw_MConst) = Format(.êîó , "#,###")       '< ïKóvêî >
                        End Select
                    Else
'                       [ ïiî‘íPà Å@åéó  ]
                        If .DD = 70 Then B1lb_NJSuT = Format(.êîó , "#,###")                '< ìñåéïiî‘êîó  >
                        If .DD = 71 Then
                            B1lb_ZNJSuT = Format(.êîó , "#,###")                            '< ëOåéïiî‘êîó  >
                            B1lb_SaNJSuT = Format(CisFun.Val2(B1lb_NJSuT) - .êîó , "#,###") '< ç∑ >
                        End If
                    End If
                End If
            End With

            Call HNJReadNext
        Loop
    End If
    Call HNJClose

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
'* ïiî‘èÓïÒ
'*---------*
    Call ItemsClearHTM
    
    gSL_Select = "Select HT.*," & vbCrLf
    gSL_Select = gSL_Select & " IsNull( HN.ç›å…ä«óùãÊï™, 0 ) ç›å…ä«óùãÊï™," & vbCrLf
    gSL_Select = gSL_Select & " IsNull( HN.ç›å…ä«óùêîíl, 0 ) ç›å…ä«óùêîíl," & vbCrLf
    gSL_Select = gSL_Select & " IsNull( SY.ílñºèÃ, '' ) ç›å…ä«óùñº" & vbCrLf
    gSL_Select = gSL_Select & " From ïiî‘éÊà¯êÊÉ}ÉXÉ^ HT" & vbCrLf
    gSL_Select = gSL_Select & " Left Outer Join ïiî‘É}ÉXÉ^ HN" & vbCrLf
    gSL_Select = gSL_Select & " On  HN.ïiî‘ = HT.ïiî‘" & vbCrLf
    gSL_Select = gSL_Select & " Left Outer Join ñºèÃÉ}ÉXÉ^ SY" & vbCrLf
    gSL_Select = gSL_Select & " On  SY.ãÊï™ñºèÃ = 'ç›å…ãÊï™'" & vbCrLf
    gSL_Select = gSL_Select & " And SY.íl = HN.ç›å…ä«óùãÊï™" & vbCrLf
    gSL_Select = gSL_Select & " And SY.ãÊï™É^ÉCÉv = ''" & vbCrLf
    gSL_Select = gSL_Select & " Where HT.ïiî‘   = '" & RTrim(iH1_Hinbn) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   HT.éÊà¯êÊ = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   HT.éÛì¸   = '" & RTrim(iH1_Ukeir) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   HT.ïiñ⁄   = 1" & vbCrLf
    If HTMRead(gSL_Select, 1) Then
        With HTM
            
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
            If .ì‡é¶î≠íç = 0 Then B1lb_NaHaKB = "ÇµÇ»Ç¢" Else B1lb_NaHaKB = "Ç∑ÇÈ"
            NaiHatu_Back(2) = B1lb_NaHaKB
        
'           [ î‰ó¶ ]
            B1lb_Ritu = Format(.î≠íçî‰ó¶, "#Åì;#;#")
    
'           [ óLå¯ä˙å¿ ]
            B1lb_SDate = Format(RTrim$(.óLå¯ä˙å¿äJén), "@@@@/@@/@@")
            B1lb_EDate = Format(RTrim$(.óLå¯ä˙å¿èIóπ), "@@@@/@@/@@")
        
'           [ íºëóêÊ ]
            If .éËîzãÊï™ = 1 Then sv_Choku = RTrim$(.î[ì¸êÊ)
        
        End With
    End If

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/     ëOÅ@åéÅ@ññÅ@ç›Å@å…Å@éZÅ@èo
'_/
    With CisDB
        .SQL = "ç›å…êîéZèo"
        .StoadoCount = 5
        .ParaNo = 0:    .ParaIO = Returnóp
        .ParaNo = 1:    .ParaIO = Inputóp
        .ParaNo = 2:    .ParaIO = Inputóp
        .ParaNo = 3:    .ParaIO = Inputóp
        .ParaNo = 4:    .ParaIO = OutPutóp
        .ParaNo = 5:    .ParaIO = OutPutóp
        .ParaNo = 1:    .ParaValue = RTrim$(iH1_YM)
        .ParaNo = 2:    .ParaValue = RTrim$(iH1_Hinbn)
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
    
    If HTM.SKãÊï™ = 1 And HTM.ì‡é¶î≠íç = 1 Then

        gSL_Select = "Select Sum( î≠íçêî ) î≠íçêî"
        gSL_Select = gSL_Select & " From î≠íçÉeÅ[ÉuÉã"
        gSL_Select = gSL_Select & " Where ïiî‘   = '" & RTrim$(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " And   édì¸êÊ = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & " And   éÛì¸   = '" & RTrim(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & " And   SubString( î[ì¸ì˙, 1, 6 ) = '" & RTrim$(iH1_YM) & "'"
        If HCTRead(gSL_Select, 1) Then B1lb_HCSu = Format(HCT.î≠íçêî, "#,###")
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
        .SQL = "DELETE  î≠íçì‡é¶ÉeÅ[ÉuÉã "
        .SQL = .SQL & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        .SQL = .SQL & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
        .SQL = .SQL & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        .SQL = .SQL & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
        .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
        .DBExec
        
        .SQL = "DELETE  î≠íçì‡é¶ëOâÒÉeÅ[ÉuÉã "
        .SQL = .SQL & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        .SQL = .SQL & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
        .SQL = .SQL & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
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
    Dim MntFlg          As Boolean

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
    
' *ADD*
    MntFlg = False
    For gInt = 1 To 31
        If CisFun.Val2(iB1_Su(gInt)) <> CisFun.Val2(iB1_Su(gInt).Tag) Then
            MntFlg = True
            Exit For
        End If
    Next gInt
    If MntFlg And SyoriKB = "MNT" Then
    ' *ëOâÒì‡óe çÌèú*
        With CisDB
            .SQL = "DELETE î≠íçì‡é¶ëOâÒÉeÅ[ÉuÉã "
            .SQL = .SQL & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
            .SQL = .SQL & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
            .SQL = .SQL & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
            .SQL = .SQL & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
            .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
            .DBExec
        End With
    ' *ëOâÒÉfÅ[É^ çÏê¨*
        With CisDB
            .SQL = "INSERT INTO î≠íçì‡é¶ëOâÒÉeÅ[ÉuÉã "
            .SQL = .SQL & " (ì‡é¶îNåé, ïiî‘, éËîzêÊ, éÛì¸, íºëóêÊ, "
            .SQL = .SQL & "  î[ä˙, édä|ì˙, éËîzì˙, êîó , "
            .SQL = .SQL & "  èàóùãÊï™, çÌèúãÊï™, ïœçXâÒêî, ÉçÅ[ÉåÉxÉã, åvâÊîNåé, "
            .SQL = .SQL & "  çÏê¨é“, çÏê¨ì˙, çÏê¨í[ññ )"
            .SQL = .SQL & "  SELECT ì‡é¶îNåé, ïiî‘, éËîzêÊ, éÛì¸, íºëóêÊ, "
            .SQL = .SQL & "         î[ä˙, édä|ì˙, éËîzì˙, êîó , "
            .SQL = .SQL & "         èàóùãÊï™, çÌèúãÊï™, ïœçXâÒêî, ÉçÅ[ÉåÉxÉã, åvâÊîNåé, "
            .SQL = .SQL & "'" & gTanto & "',"
            .SQL = .SQL & "GETDATE(),HOST_NAME() "
            .SQL = .SQL & "         FROM î≠íçì‡é¶ÉeÅ[ÉuÉã "
            .SQL = .SQL & "         Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
            .SQL = .SQL & "           And åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
            .SQL = .SQL & "           And ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
            .SQL = .SQL & "           And éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
            .SQL = .SQL & "           And éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
            .DBExec
        End With
    End If
    
    ClwFlg = False
    For gInt = 1 To 31
        If CisFun.Val2(iB1_Su(gInt)) <> CisFun.Val2(iB1_Su(gInt).Tag) Then
            
            gStr = RTrim(iH1_YM) & Format(gInt, "00")
            gSL_Select = "SELECT * FROM î≠íçì‡é¶ÉeÅ[ÉuÉã "
            gSL_Select = gSL_Select & " WHERE ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
            gSL_Select = gSL_Select & "   AND ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
            gSL_Select = gSL_Select & "   AND éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
            gSL_Select = gSL_Select & "   AND éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
            gSL_Select = gSL_Select & "   AND î[ä˙ = '" & RTrim(gStr) & "'"
            If HNJRead(gSL_Select, 1) Then
                AddFlg = False
' *DEL* 2004/04/09
'                With HNJ
'                    HNZ.ì‡é¶îNåé = .ì‡é¶îNåé
'                    HNZ.ïiî‘ = .ïiî‘
'                    HNZ.éËîzêÊ = .éËîzêÊ
'                    HNZ.éÛì¸ = .éÛì¸
'                    HNZ.íºëóêÊ = .íºëóêÊ
'                    HNZ.î[ä˙ = .î[ä˙
'                    HNZ.édä|ì˙ = .édä|ì˙
'                    HNZ.éËîzì˙ = .éËîzì˙
'                    HNZ.êîó  = .êîó 
'                    HNZ.èàóùãÊï™ = .èàóùãÊï™
'                    HNZ.çÌèúãÊï™ = .çÌèúãÊï™
'                    HNZ.ïœçXâÒêî = 1
'                    HNZ.ÉçÅ[ÉåÉxÉã = .ÉçÅ[ÉåÉxÉã
'                    HNZ.åvâÊîNåé = .åvâÊîNåé
'                End With
'
'                If RTrim(B1lb_LSu(gInt).Tag) = "" Then
'                    HNZ.çÏê¨é“ = gTanto
'                    Call HNZInsert
'                Else
'                    HNZ.çXêVé“ = gTanto
'                    Call HNZUpdate
'                End If
            Else
                
            ' édä|ÅEéËîzì˙
                With CisDB
                    .SQL = "ÉÅÉìÉeópédä|éËîzì˙éZèo"
                    .StoadoCount = 6
                    .ParaNo = 0:    .ParaIO = Returnóp
                    .ParaNo = 1:    .ParaIO = Inputóp
                    .ParaNo = 2:    .ParaIO = Inputóp
                    .ParaNo = 3:    .ParaIO = Inputóp
                    .ParaNo = 4:    .ParaIO = Inputóp
                    .ParaNo = 5:    .ParaIO = OutPutóp
                    .ParaNo = 6:    .ParaIO = OutPutóp
                    
                    .ParaNo = 1:    .ParaValue = RTrim(iH1_Hinbn)
                    .ParaNo = 2:    .ParaValue = RTrim(iH1_Torcd)
                    .ParaNo = 3:    .ParaValue = RTrim(iH1_Ukeir)
                    .ParaNo = 4:    .ParaValue = RTrim(gStr)
                    
                    .DBStored
                
                    .ParaNo = 0
                    If .ParaValue <> 0 Then
                        CisFun.MB_Lines = 5
                        CisFun.MB_MSG(2) = "    ÉJÉåÉìÉ_ÉèÅ[ÉNçÏê¨Ç…é∏îsÇµÇ‹ÇµÇΩÅB    "
                        CisFun.MB_Button = OK
                        CisFun.MBOX
                        Exit Function
                    End If
                
                    .ParaNo = 5:    wsSYMD = .ParaValue
                    .ParaNo = 6:    wsTYMD = .ParaValue
                End With
                
                AddFlg = True
                Call ItemsClearHNJ
                With HNJ
                    .ì‡é¶îNåé = RTrim(iH1_NJYM)
                    .ïiî‘ = RTrim(iH1_Hinbn)
                    .éËîzêÊ = RTrim(iH1_Torcd)
                    .éÛì¸ = RTrim(iH1_Ukeir)
                    .íºëóêÊ = RTrim$(sv_Choku)
                    .î[ä˙ = RTrim(gStr)
                    .édä|ì˙ = RTrim(wsSYMD)
                    .éËîzì˙ = RTrim(wsTYMD)
                End With
            End If
    
            HNJ.èàóùãÊï™ = 1
            HNJ.êîó  = CisFun.Val2(iB1_Su(gInt))
            HNJ.åvâÊîNåé = RTrim(iH1_YM)
            If HNJ.êîó  <> 0 Then
                If AddFlg Then
                    HNJ.çÏê¨é“ = gTanto
                    Call HNJInsert
                    iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
                Else
                    HNJ.çXêVé“ = gTanto
                    Call HNJUpdate
                    iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
                End If
            Else
        ' ( çÌèúèàóù )
        
                gSL_Select = "Insert Into î≠íçì‡é¶çÌèúÉèÅ[ÉN                "
                gSL_Select = gSL_Select & " Select ì‡é¶îNåé, ïiî‘, éËîzêÊ, éÛì¸, íºëóêÊ,"
                gSL_Select = gSL_Select & " î[ä˙, édä|ì˙, éËîzì˙, êîó , èàóùãÊï™, çÌèúãÊï™,"
                gSL_Select = gSL_Select & " ïœçXâÒêî, ÉçÅ[ÉåÉxÉã, åvâÊîNåé,"
                gSL_Select = gSL_Select & " GETDATE() çÏê¨ì˙, SubString( Host_Name(), 1, 20 ) çÏê¨í[ññ"
                gSL_Select = gSL_Select & " From î≠íçì‡é¶ÉeÅ[ÉuÉã "
                gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
                gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
                gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
                gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
                gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
                gSL_Select = gSL_Select & " And   î[ä˙ = '" & RTrim(gStr) & "'"
                With CisDB
                    .SQL = gSL_Select
                    .DBExec
                End With
                
                With CisDB
                    .SQL = "DELETE î≠íçì‡é¶ÉeÅ[ÉuÉã "
                    .SQL = .SQL & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
                    .SQL = .SQL & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
                    .SQL = .SQL & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
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
        gSL_Select = gSL_Select & " Where ãÊï™ = 0"
        gSL_Select = gSL_Select & " And   î≠íçîNåé = '" & RTrim$(iH1_YM) & "'"
        gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim$(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim$(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim$(iH1_Hinbn) & "'"
        If NHCRead(gSL_Select, 1) Then
            With CisDB
                .SQL = "Delete ì‡é¶î≠íçÉ`ÉFÉbÉNÉeÅ[ÉuÉã "
                .SQL = .SQL & " Where ãÊï™ = 0"
                .SQL = .SQL & " And   î≠íçîNåé = '" & RTrim$(iH1_YM) & "'"
                .SQL = .SQL & " And   éËîzêÊ = '" & RTrim$(iH1_Torcd) & "'"
                .SQL = .SQL & " And   éÛì¸ = '" & RTrim$(iH1_Ukeir) & "'"
                .SQL = .SQL & " And   ïiî‘ = '" & RTrim$(iH1_Hinbn) & "'"
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
        If RTrim$(iH1_Hinbn) = "" Then Exit Function

'       < í«â¡éûÅFïiî‘éÊà¯êÊÉ}ÉXÉ^ >
        gSL_Select = "Select éÊà¯êÊ, éÛì¸"
        gSL_Select = gSL_Select & " From ïiî‘éÊà¯êÊÉ}ÉXÉ^"
        gSL_Select = gSL_Select & " Where ïiî‘ = '" & RTrim$(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " And   ïiñ⁄ = 1"
        If HTMRead(gSL_Select) Then
            If CisDB.RecordCount > 1 Then
                Call HTMClose
                Call Look_Tehacd
                sv_Hinbn = RTrim$(iH1_Hinbn)
            Else
                iH1_Torcd = HTM.éÊà¯êÊ
                iH1_Ukeir = HTM.éÛì¸
                Call HTMClose
            End If
        End If
    
    Else
        If RTrim$(iH1_NJYM) = "" Or _
           RTrim$(iH1_YM) = "" Or _
           RTrim$(iH1_Hinbn) = "" Then Exit Function

'       * î≠íçì‡é¶ÉeÅ[ÉuÉã ë∂ç›É`ÉFÉbÉN
        gSL_Select = "SELECT éËîzêÊ, éÛì¸"
        gSL_Select = gSL_Select & " FROM î≠íçì‡é¶ÉeÅ[ÉuÉã "
        gSL_Select = gSL_Select & " WHERE ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        gSL_Select = gSL_Select & "   AND åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
        gSL_Select = gSL_Select & "   AND ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " GROUP BY éËîzêÊ, éÛì¸"
        If HNJRead(gSL_Select) Then
        ' * éËîzêÊëIë
            If CisDB.RecordCount > 1 Then
                Call HNJClose
                Call Look_Tehacd
                sv_Hinbn = RTrim$(iH1_Hinbn)
            Else
                iH1_Torcd = HNJ.éËîzêÊ
                iH1_Ukeir = HNJ.éÛì¸
                Call HNJClose
            End If
        End If
    End If

SetTorcd_ED:
    SetTorcd = True
End Function
'Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°
'Å†
'Å†         å„Å@ï˚Å@èàÅ@óùÅ@( í èÌÉÅÉìÉeéû )
'Å†
'Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°Å†Å°
Private Sub After_Proc()

'   < èàóùëŒè€ÉfÅ[É^É`ÉFÉbÉN >
    gSL_Select = "IF EXISTS("
'Å§Å§-------------------------------------------------------<< 2005/03/08 >> Update Start
'   gSL_Select = gSL_Select & " SELECT ïiî‘ FROM î≠íçì‡é¶ÉeÅ[ÉuÉã"
'   gSL_Select = gSL_Select & " Where èàóùãÊï™ <> 0"
'   gSL_Select = gSL_Select & " And   CASE WHEN ISNULL( çXêVí[ññ, '' ) = '' THEN çÏê¨í[ññ ELSE çXêVí[ññ END"
'   gSL_Select = gSL_Select & "     = SUBSTRING( HOST_NAME(), 1, 20 )"
'   gSL_Select = gSL_Select & " And   SUBSTRING( î[ä˙, 1, 6 ) = '" & gn_HNYM & "'"
'   gSL_Select = gSL_Select & " Union"
'   gSL_Select = gSL_Select & " Select ïiî‘ From î≠íçì‡é¶çÌèúÉèÅ[ÉN"
'   gSL_Select = gSL_Select & " Where çÏê¨í[ññ = SUBSTRING( HOST_NAME(), 1, 20 )"
'   ì‡é¶îNåéÅïåvâÊîNåé
    gSL_Select = gSL_Select & " select ïiî‘ from î≠íçì‡é¶ÉeÅ[ÉuÉã"
    gSL_Select = gSL_Select & "  where èàóùãÊï™ <> 0"
'Å§Å§-------------------------------------------------------<< 2005/03/19 >> Update Start
   'gSL_Select = gSL_Select & "    and case when çXêVí[ññ = '' then çÏê¨í[ññ else çXêVí[ññ end = substring(host_name(), 1, 20)"
    gSL_Select = gSL_Select & "    and case when isnull(çXêVí[ññ, '') = '' then çÏê¨í[ññ else çXêVí[ññ end = substring(host_name(), 1, 20)"
'Å¢Å¢-------------------------------------------------------<< 2005/03/19 >> Update End
    gSL_Select = gSL_Select & "    and ì‡é¶îNåé = '" & gn_HNYM & "'"
    gSL_Select = gSL_Select & "    and åvâÊîNåé = '" & gn_HNYM & "'"
    gSL_Select = gSL_Select & " union"
    gSL_Select = gSL_Select & " select ïiî‘ from î≠íçì‡é¶çÌèúÉèÅ[ÉN"
    gSL_Select = gSL_Select & "  where çÏê¨í[ññ = substring(host_name(), 1, 20)"
    gSL_Select = gSL_Select & "    and ì‡é¶îNåé = '" & gn_HNYM & "'"
    gSL_Select = gSL_Select & "    and åvâÊîNåé = '" & gn_HNYM & "'"
'Å¢Å¢-------------------------------------------------------<< 2005/03/08 >> Update End
    gSL_Select = gSL_Select & " ) BEGIN SELECT 1 åèêî END"
    gSL_Select = gSL_Select & " ELSE"
    gSL_Select = gSL_Select & "   BEGIN SELECT 0 åèêî END"
    Call HNJRead(gSL_Select, 1)
'Å§Å§-------------------------------------------------------<< 2005/03/08 >> Update Start
'   If HNJ.åèêî = 0 Then Exit Sub
    If HNJ.åèêî = 0 Then GoTo After_Proc_Clear
'Å¢Å¢-------------------------------------------------------<< 2005/03/08 >> Update End
    
    Erase g315_InsCnt:  Erase g315_UpdCnt:  g315_ErrCnt = 0

    g315_YM = gn_HNYM
    g315_ProcKB = 1
    g315_FProc = 1

    CNJ0315.Show vbModal

    Unload CNJ0315
    Set CNJ0315 = Nothing

'   <ïΩèÄâªê›íËçXêVÉGÉâÅ[ÉäÉXÉg>
    Call HJKErr_List

'Å§Å§-------------------------------------------------------<< 2005/03/08 >> Insert Start
After_Proc_Clear:
'Å¢Å¢-------------------------------------------------------<< 2005/03/08 >> Insert End
'   < èàóùãÊï™ÉNÉäÉA >
    gSL_Select = "Delete î≠íçì‡é¶çÌèúÉèÅ[ÉN"
    gSL_Select = gSL_Select & " Where çÏê¨í[ññ = SUBSTRING( HOST_NAME(), 1, 20 )"
    With CisDB
        .SQL = gSL_Select
        .DBExec
    End With
    
    gSL_Select = "UpDate î≠íçì‡é¶ÉeÅ[ÉuÉã Set"
    gSL_Select = gSL_Select & " èàóùãÊï™ = 0"
    gSL_Select = gSL_Select & " Where èàóùãÊï™ <> 0"
    gSL_Select = gSL_Select & " And   CASE WHEN ISNULL( çXêVí[ññ, '' ) = '' THEN çÏê¨í[ññ ELSE çXêVí[ññ END"
    gSL_Select = gSL_Select & "     = SUBSTRING( HOST_NAME(), 1, 20 )"
    With CisDB
        .SQL = gSL_Select
        .DBExec
    End With

End Sub
