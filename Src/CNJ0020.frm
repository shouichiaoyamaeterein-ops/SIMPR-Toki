VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Begin VB.Form CNJ0020 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "ê∂éYåvâÊì¸óÕ"
   ClientHeight    =   11235
   ClientLeft      =   2220
   ClientTop       =   1830
   ClientWidth     =   15360
   ControlBox      =   0   'False
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
   ForeColor       =   &H00C0C000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'Z µ∞¿ﬁ∞
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   WindowState     =   2  'ç≈ëÂâª
   Begin VB.PictureBox B1_Area2 
      Height          =   1515
      Left            =   720
      ScaleHeight     =   1455
      ScaleWidth      =   13875
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   9120
      Width           =   13935
      Begin Cis3D_v60.CIS3D CIS3D31 
         Height          =   1395
         Left            =   30
         Top             =   30
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   2461
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
         Begin Cis3D_v60.CIS3D CIS3D63 
            Height          =   525
            Left            =   8490
            Top             =   840
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   926
            ForeColor       =   16711680
            Caption         =   "ëOåéì‡é¶"
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
            Begin Cis3D_v60.CIS3D B2lb_GZNJSu 
               Height          =   360
               Left            =   1170
               Top             =   90
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   635
               BackColor       =   16777152
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
            Begin Cis3D_v60.CIS3D B2lb_Kesu3 
               Height          =   330
               Left            =   90
               Top             =   390
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   582
               BackColor       =   16777152
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
            Begin Cis3D_v60.CIS3D B2lb_Kesu2 
               Height          =   330
               Left            =   90
               Top             =   390
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   582
               BackColor       =   16777152
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
            Height          =   525
            Left            =   11190
            Top             =   840
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   926
            ForeColor       =   16711680
            Caption         =   "ì‡é¶ç∑"
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
            Begin Cis3D_v60.CIS3D B2lb_GSNJSu 
               Height          =   360
               Left            =   1080
               Top             =   90
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   635
               BackColor       =   16777152
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
            Height          =   525
            Left            =   1380
            Top             =   840
            Width           =   4380
            _ExtentX        =   7726
            _ExtentY        =   926
            ForeColor       =   16711680
            Caption         =   "óLå¯ä˙å¿"
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
            cPositionX      =   45
            Begin Cis3D_v60.CIS3D B2lb_SDate 
               Height          =   360
               Left            =   1170
               Top             =   90
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cBoderStyle     =   1
            End
            Begin Cis3D_v60.CIS3D B2lb_EDate 
               Height          =   360
               Left            =   2880
               Top             =   90
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cBoderStyle     =   1
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   300
               Index           =   41
               Left            =   2550
               Top             =   150
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
            Begin Cis3D_v60.CIS3D B2lb_Zaiko 
               Height          =   330
               Left            =   120
               Top             =   390
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   582
               BackColor       =   16777152
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
            Begin Cis3D_v60.CIS3D B2lb_Kesu1 
               Height          =   330
               Left            =   90
               Top             =   390
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   582
               BackColor       =   16777152
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
            Height          =   525
            Left            =   30
            Top             =   840
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   926
            ForeColor       =   16711680
            Caption         =   "î‰ó¶"
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
            cPositionX      =   60
            Begin Cis3D_v60.CIS3D B2lb_Ritu 
               Height          =   360
               Left            =   630
               Top             =   90
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   635
               BackColor       =   16777152
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
            Height          =   525
            Left            =   5760
            Top             =   840
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   926
            ForeColor       =   16711680
            Caption         =   "ìñåéì‡é¶"
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
            Begin Cis3D_v60.CIS3D B2lb_GNJSu 
               Height          =   360
               Left            =   1200
               Top             =   90
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   635
               BackColor       =   16777152
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
            Begin Cis3D_v60.CIS3D B2lb_Cycle 
               Height          =   330
               Left            =   90
               Top             =   390
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   582
               BackColor       =   16777152
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
            Begin Cis3D_v60.CIS3D B2lb_Syuyo1 
               Height          =   330
               Left            =   90
               Top             =   390
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   582
               BackColor       =   16777152
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
            Begin Cis3D_v60.CIS3D B2lb_CDate 
               Height          =   330
               Left            =   90
               Top             =   390
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   582
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cBoderStyle     =   1
            End
            Begin Cis3D_v60.CIS3D B2lb_CBin 
               Height          =   330
               Left            =   1380
               Top             =   390
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   582
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
            Width           =   2595
            _ExtentX        =   4577
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
            Begin Cis3D_v60.CIS3D B2lb_ZKKbn 
               Height          =   330
               Left            =   120
               Top             =   390
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   582
               BackColor       =   16777152
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
            Begin Cis3D_v60.CIS3D B2lb_ZKSu 
               Height          =   330
               Left            =   1320
               Top             =   390
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   582
               BackColor       =   16777152
               Caption         =   "000,000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Begin Cis3D_v60.CIS3D B2lb_Syuyo2 
               Height          =   330
               Left            =   90
               Top             =   390
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   582
               BackColor       =   16777152
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
         Begin Cis3D_v60.CIS3D CIS3D12 
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
            Begin Cis3D_v60.CIS3D B2lb_SKKbn 
               Height          =   330
               Left            =   90
               Top             =   390
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   582
               BackColor       =   16777152
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
      Height          =   1560
      Left            =   750
      ScaleHeight     =   1500
      ScaleWidth      =   9375
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   765
      Width           =   9435
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   855
         Left            =   75
         Top             =   570
         Width           =   1470
         _ExtentX        =   2593
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
         cPositionY      =   30
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   375
            Left            =   60
            TabIndex        =   0
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
         Left            =   1545
         Top             =   570
         Width           =   3270
         _ExtentX        =   5768
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
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   60
            TabIndex        =   1
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
         Left            =   4815
         Top             =   570
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   1508
         ForeColor       =   16711680
         Caption         =   "éË    îz    êÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cPositionX      =   360
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D H1lb_Torcd 
            Height          =   375
            Left            =   1050
            Top             =   390
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   661
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
            Left            =   3210
            Top             =   30
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
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   60
            TabIndex        =   2
            Top             =   390
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
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   375
            Left            =   3300
            TabIndex        =   3
            Top             =   390
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
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   855
         Left            =   8655
         Top             =   570
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1508
         ForeColor       =   16711680
         Caption         =   "çHãÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D H1lb_KOKbn 
            Height          =   360
            Left            =   135
            Top             =   375
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "X"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   495
         Left            =   75
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
         cPositionX      =   180
         Begin CisYM_V60.CisYM iH1_NJYM 
            Height          =   375
            Left            =   1350
            TabIndex        =   4
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
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   495
         Index           =   4
         Left            =   2835
         Top             =   75
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   873
         ForeColor       =   32768
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
      Left            =   10365
      ScaleHeight     =   1275
      ScaleWidth      =   4215
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   990
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
         Top             =   630
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
         Top             =   630
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
         Top             =   630
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
         Top             =   630
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
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   37
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
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "Åy ê∂éYåvâÊì¸óÕ  Åz"
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
         Left            =   14220
         Top             =   60
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
      Height          =   6240
      Left            =   735
      ScaleHeight     =   6180
      ScaleWidth      =   13860
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2700
      Width           =   13920
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   6075
         Left            =   60
         Top             =   60
         Width           =   13755
         _ExtentX        =   24262
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
            Width           =   13695
            _ExtentX        =   24156
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
               Height          =   1470
               Index           =   30
               Left            =   1080
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   72
               TabStop         =   0   'False
               Top             =   4515
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   30
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇRÇPì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   30
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   30
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   29
               Left            =   12420
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   29
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇRÇOì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   29
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   29
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   28
               Left            =   11160
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   70
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   28
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇXì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   28
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   28
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   27
               Left            =   9900
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   27
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇWì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   27
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   27
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   26
               Left            =   8640
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   68
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   26
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇVì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   26
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   26
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   25
               Left            =   7380
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   67
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   25
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇUì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   25
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   25
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   24
               Left            =   6120
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   24
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇTì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   24
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   24
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   23
               Left            =   4860
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   23
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇSì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   23
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   23
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   22
               Left            =   3600
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   22
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇRì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   22
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   22
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   21
               Left            =   2340
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   21
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇQì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   21
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   21
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   20
               Left            =   1080
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   3015
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   20
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇPì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   20
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   20
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   19
               Left            =   12420
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   19
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQÇOì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   19
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   19
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   18
               Left            =   11160
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   18
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇXì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   18
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   18
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   17
               Left            =   9900
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   17
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇWì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   17
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   17
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   16
               Left            =   8640
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   16
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇVì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   16
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   16
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   15
               Left            =   7380
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   15
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇUì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   15
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   15
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   14
               Left            =   6120
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   14
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇTì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   14
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   14
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   13
               Left            =   4860
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   13
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇSì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   13
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   13
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   12
               Left            =   3600
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   12
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇRì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   12
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   12
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   11
               Left            =   2340
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   11
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇQì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   11
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   11
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   10
               Left            =   1080
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   1515
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   10
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇPì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   10
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   10
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   9
               Left            =   12420
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   15
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   9
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇPÇOì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   9
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   9
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   8
               Left            =   11160
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   15
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   8
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇXì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   8
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   8
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   7
               Left            =   9900
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   15
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   7
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇWì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   7
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   7
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   6
               Left            =   8640
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   15
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   6
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇVì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   6
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   6
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   5
               Left            =   7380
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   15
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   5
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇUì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   5
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   5
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   4
               Left            =   6120
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   15
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   4
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇTì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   4
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   4
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   3
               Left            =   4860
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   15
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   3
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇSì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   3
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   3
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   2
               Left            =   3600
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   15
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   2
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇRì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   2
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   2
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   1
               Left            =   2340
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   15
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   1
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  ForeColor       =   16711680
                  Caption         =   "ÇQì˙ "
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
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   1
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   1
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   1455
               Index           =   48
               Left            =   7800
               Top             =   4530
               Width           =   5850
               _ExtentX        =   10319
               _ExtentY        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_LTime 
                  Height          =   345
                  Left            =   1740
                  Top             =   210
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   609
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
                  cPositionX      =   90
                  cPositionY      =   50
               End
               Begin Cis3D_v60.CIS3D CIS3D42 
                  Height          =   375
                  Left            =   750
                  Top             =   210
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   661
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
                  Height          =   405
                  Left            =   2640
                  Top             =   195
                  Width           =   3015
                  _ExtentX        =   5318
                  _ExtentY        =   714
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
                     Height          =   345
                     Index           =   1
                     Left            =   1110
                     Top             =   30
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   609
                     BackColor       =   16777152
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
                     Height          =   345
                     Index           =   2
                     Left            =   1470
                     Top             =   30
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   609
                     BackColor       =   16777152
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
                     Height          =   345
                     Index           =   3
                     Left            =   1830
                     Top             =   30
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   609
                     BackColor       =   16777152
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
                     Height          =   345
                     Index           =   4
                     Left            =   2190
                     Top             =   30
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   609
                     BackColor       =   16777152
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
                     Height          =   345
                     Index           =   5
                     Left            =   2550
                     Top             =   30
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   609
                     BackColor       =   16777152
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
                  Height          =   405
                  Left            =   780
                  Top             =   900
                  Width           =   2025
                  _ExtentX        =   3572
                  _ExtentY        =   714
                  ForeColor       =   16711680
                  Caption         =   "ê∂éY€Øƒ"
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
                     Height          =   345
                     Left            =   960
                     Top             =   0
                     Width           =   975
                     _ExtentX        =   1720
                     _ExtentY        =   609
                     BackColor       =   16777152
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
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -20
                  End
               End
               Begin Cis3D_v60.CIS3D CIS3D45 
                  Height          =   375
                  Left            =   3150
                  Top             =   900
                  Width           =   2505
                  _ExtentX        =   4419
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
                     Height          =   345
                     Index           =   1
                     Left            =   600
                     Top             =   0
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   609
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
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -20
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Day 
                     Height          =   345
                     Index           =   2
                     Left            =   960
                     Top             =   0
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   609
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
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -20
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Day 
                     Height          =   345
                     Index           =   3
                     Left            =   1320
                     Top             =   0
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   609
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
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -20
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Day 
                     Height          =   345
                     Index           =   4
                     Left            =   1680
                     Top             =   0
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   609
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
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -20
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Day 
                     Height          =   345
                     Index           =   5
                     Left            =   2040
                     Top             =   0
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   609
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
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -20
                  End
               End
               Begin VB.Shape Shape4 
                  Height          =   645
                  Left            =   540
                  Shape           =   4  'ä€Ç›ÇÃÇ†ÇÈí∑ï˚å`
                  Top             =   750
                  Width           =   2415
               End
               Begin VB.Shape Shape3 
                  Height          =   645
                  Left            =   2970
                  Shape           =   4  'ä€Ç›ÇÃÇ†ÇÈí∑ï˚å`
                  Top             =   750
                  Width           =   2805
               End
               Begin VB.Shape Shape2 
                  Height          =   645
                  Left            =   2460
                  Shape           =   4  'ä€Ç›ÇÃÇ†ÇÈí∑ï˚å`
                  Top             =   60
                  Width           =   3315
               End
               Begin VB.Shape Shape1 
                  Height          =   645
                  Left            =   540
                  Shape           =   4  'ä€Ç›ÇÃÇ†ÇÈí∑ï˚å`
                  Top             =   60
                  Width           =   1905
               End
            End
            Begin Cis3D_v60.CIS3D B1lb_NYM 
               Height          =   315
               Index           =   2
               Left            =   5940
               Top             =   4500
               Width           =   1830
               _ExtentX        =   3228
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
            Begin Cis3D_v60.CIS3D B1lb_GZNJSu 
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
               Begin Cis3D_v60.CIS3D B1lb_GSKSu 
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
               Begin Cis3D_v60.CIS3D B1lb_GSKSu 
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
            Begin Cis3D_v60.CIS3D B1lb_GZNJSu 
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
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   585
               Index           =   0
               Left            =   60
               Top             =   900
               Width           =   255
               _ExtentX        =   450
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
            End
            Begin Cis3D_v60.CIS3D CIS3D3 
               Height          =   585
               Index           =   0
               Left            =   60
               Top             =   300
               Width           =   255
               _ExtentX        =   450
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
            End
            Begin VB.PictureBox B_Sur 
               BackColor       =   &H00808080&
               BorderStyle     =   0  'Ç»Çµ
               Height          =   1470
               Index           =   0
               Left            =   1080
               ScaleHeight     =   1470
               ScaleWidth      =   1260
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   15
               Width           =   1260
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   0
                  Left            =   -15
                  TabIndex        =   5
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
               Begin Cis3D_v60.CIS3D BLB_Day 
                  Height          =   300
                  Index           =   0
                  Left            =   -15
                  Top             =   -15
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  BackColor       =   8421631
                  ForeColor       =   16711680
                  Caption         =   "ÇPì˙ "
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
                     Index           =   0
                     Left            =   60
                     Top             =   60
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   370
                     BackColor       =   8421631
                     ForeColor       =   0
                     Caption         =   "íã"
                     BackStyle       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Index           =   0
                  Left            =   0
                  Top             =   1170
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZNJSu 
                  Height          =   300
                  Index           =   0
                  Left            =   0
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  BackColor       =   16761087
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
                  cPositionX      =   -15
               End
               Begin Cis3D_v60.CIS3D B1lb_ZSKSu 
                  Height          =   300
                  Index           =   0
                  Left            =   0
                  Top             =   570
                  Width           =   1260
                  _ExtentX        =   2223
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
                  cPositionX      =   -15
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DDW 
               Height          =   300
               Index           =   0
               Left            =   30
               Top             =   0
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   529
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
            Begin Cis3D_v60.CIS3D BLB_ 
               Height          =   285
               Index           =   2
               Left            =   330
               Top             =   300
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "ì‡é¶êî"
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
               cAlingnment     =   6
            End
            Begin Cis3D_v60.CIS3D BLB_ 
               Height          =   285
               Index           =   5
               Left            =   330
               Top             =   1200
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "ì‡é¶êî"
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
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_ 
               Height          =   645
               Index           =   50
               Left            =   330
               Top             =   570
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   1138
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   585
               Index           =   1
               Left            =   60
               Top             =   2400
               Width           =   255
               _ExtentX        =   450
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
            End
            Begin Cis3D_v60.CIS3D CIS3D3 
               Height          =   585
               Index           =   1
               Left            =   60
               Top             =   1800
               Width           =   255
               _ExtentX        =   450
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
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   300
               Index           =   14
               Left            =   30
               Top             =   1500
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   529
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
            Begin Cis3D_v60.CIS3D BLB_ 
               Height          =   285
               Index           =   15
               Left            =   330
               Top             =   1800
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "ì‡é¶êî"
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
               cFont3D         =   2
               cAlingnment     =   6
            End
            Begin Cis3D_v60.CIS3D BLB_ 
               Height          =   285
               Index           =   16
               Left            =   330
               Top             =   2700
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "ì‡é¶êî"
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
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_ 
               Height          =   645
               Index           =   17
               Left            =   330
               Top             =   2070
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   1138
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   585
               Index           =   2
               Left            =   60
               Top             =   3900
               Width           =   255
               _ExtentX        =   450
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
            End
            Begin Cis3D_v60.CIS3D CIS3D3 
               Height          =   585
               Index           =   2
               Left            =   60
               Top             =   3300
               Width           =   255
               _ExtentX        =   450
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
            End
            Begin Cis3D_v60.CIS3D BLB_ 
               Height          =   300
               Index           =   28
               Left            =   30
               Top             =   3000
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   529
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
            Begin Cis3D_v60.CIS3D BLB_ 
               Height          =   285
               Index           =   29
               Left            =   330
               Top             =   3300
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "ì‡é¶êî"
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
               cFont3D         =   2
               cAlingnment     =   6
            End
            Begin Cis3D_v60.CIS3D BLB_ 
               Height          =   285
               Index           =   30
               Left            =   330
               Top             =   4200
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "ì‡é¶êî"
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
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_ 
               Height          =   645
               Index           =   31
               Left            =   330
               Top             =   3570
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   1138
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   585
               Index           =   3
               Left            =   60
               Top             =   5400
               Width           =   255
               _ExtentX        =   450
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
            End
            Begin Cis3D_v60.CIS3D CIS3D3 
               Height          =   585
               Index           =   3
               Left            =   60
               Top             =   4800
               Width           =   255
               _ExtentX        =   450
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
            End
            Begin Cis3D_v60.CIS3D BLB_ 
               Height          =   300
               Index           =   43
               Left            =   30
               Top             =   4500
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   529
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
            Begin Cis3D_v60.CIS3D BLB_ 
               Height          =   285
               Index           =   44
               Left            =   330
               Top             =   4800
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "ì‡é¶êî"
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
               cAlingnment     =   6
            End
            Begin Cis3D_v60.CIS3D BLB_ 
               Height          =   285
               Index           =   46
               Left            =   330
               Top             =   5700
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   503
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "ì‡é¶êî"
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
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_ 
               Height          =   645
               Index           =   51
               Left            =   330
               Top             =   5070
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   1138
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D B1lb_GZSKSu 
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
            Begin Cis3D_v60.CIS3D B1lb_GZSKSu 
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
               Begin Cis3D_v60.CIS3D B1lb_GSKSu 
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
            Begin Cis3D_v60.CIS3D B1lb_GZNJSu 
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
            Begin Cis3D_v60.CIS3D B1lb_GZSKSu 
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
Attribute VB_Name = "CNJ0020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :   ê∂éYåvâÊÉfÅ[É^ÉÅÉìÉeÉiÉìÉX
'**       ÉtÉHÅ[ÉÄID    :   CNJ0020
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :   2004/02/27  By CIS SIMPR-A
'**       ïœ  çX  ì˙    :   2004/04/07
'**       ïœ  çX  ì˙    :   2004/04/09  ëOâÒÉfÅ[É^çÏê¨(ïœçXï™Å®ëSÇƒ)
'**       ïœ  çX  ì˙    :   2005/03/22  ì¸óÕŒﬁ¿›Çâüâ∫éûÅAÅwé¿çséû¥◊∞'400'ÅxÇï\é¶Ç∑ÇÈéûÇ™Ç†ÇÈÅBÅiÇPïiî‘Ç≈ï°êîéËîzêÊÇ™óLÇÈéûÅj
'**       ïœ  çX  ì˙    :   2005/04/05  ëŒè€îNåéÇÃ¡™Ø∏í«â¡Å®ì‡é¶îNåéÇ…ëŒÇµÇƒÅAìñåéÅEóÇåéÅEóÇÅXåéï™ÇÃÇRÉñåéÇÃÇ›ÇnÇjÇ∆Ç∑ÇÈÅB
'**
'   200312,200312,0817233802B0,0001A,01
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' Ã´∞—±∏√®Ãﬁ Ã◊∏ﬁ
    Dim ProcHB          As String * 2   ' ì¸óÕàÊêÿë÷
    Dim SyoriKB         As String       ' èàóùãÊï™(êVãK•èCê≥•çÌèú ìô)
    Dim SyoriNM         As String       ' èàóùñºèÃ(êVãK•èCê≥•çÌèú ìô)

    Dim mOndayColor     As Long         ' â“ì≠ì˙
    Dim mOffdayColor    As Long         ' ãxì˙
    Dim mNInitColor     As Long         ' ì‡é¶
    Dim mNHenColor      As Long         ' ì‡é¶(ïœçX)
    Dim mSHenColor      As Long         ' ê∂éYåvâÊ(ïœçX)
    
    Dim wShift          As Integer      ' PageUp/PageDownéû ( ëŒè€îNåé )
    Dim sv_Hinbn        As String       ' ïiî‘(ëﬁî)
    
    Dim mNSu            As Long         ' ì˙ó êîçáåv
    Dim mNaiji          As String
    Dim wInt            As Integer

    Dim cmd_KB          As Byte
    Dim cmd_NJYM        As String
    Dim cmd_YM          As String
    Dim cmd_Hinbn       As String
    Dim cmd_Torcd       As String
    Dim cmd_Ukeir       As String
    Dim wSplit()        As String
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
        
' *ADD* 2004/04/07
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
    
    mOffdayColor = &H8080FF
    sv_Hinbn = ""
    
    mOffdayColor = BLB_Day(0).BackColor
    mOndayColor = BLB_Day(1).BackColor
    BLB_Day(0).BackColor = BLB_Day(1).BackColor
    mNHenColor = B1lb_ZNJSu(0).BackColor
    mNInitColor = B1lb_ZNJSu(1).BackColor
    B1lb_ZNJSu(0).BackColor = B1lb_ZNJSu(1).BackColor
    mSHenColor = B1lb_ZSKSu(0).BackColor
    B1lb_ZSKSu(0).BackColor = B1lb_ZSKSu(1).BackColor
    
'< ê∂éYåvâÊîNåé >
    Call GetNJProc_YM(2)
    If RTrim$(gn_SKYM) = "" Then gn_SKYM = Format(Now(), "YYYYMM")
    iH1_NJYM = gn_SKYM
    iH1_YM = gn_SKYM
    
    ' ójì˙ÉZÉbÉg
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
           Case vbKeyF1:        If PB_ADD.Visible Then Call PB_ADD_Click    'ÅyêVãKÅz
           Case vbKeyF2:        If PB_MNT.Visible Then Call PB_MNT_Click    'ÅyèCê≥Åz
           Case vbKeyF3:        If PB_DEL.Visible Then Call PB_DEL_Click    'ÅyçÌèúÅz
           Case vbKeyF5:        If PB_REP.Visible Then Call PB_REP_Click    'ÅyïœçXëOÅz
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  'ÅyåüçıÅz
           Case vbKeyPageUp:    If PB_BACK.Visible Then wShift = Shift: Call PB_Back_Click 'Åyå„ëﬁÅz
           Case vbKeyPageDown:  If PB_NEXT.Visible Then wShift = Shift: Call PB_Next_Click  'ÅyëOêiÅz
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click    'ÅyéÊè¡Åz
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    'ÅyèIóπÅz
           Case vbKeyReturn
                               If PB_ENT.Visible Then                       'Åyì¸óÕÅz
'Å§Å§---------------------------------------------<< 2005/03/22 >> Update Start
'                                 Dummy.Enabled = True: Dummy.SetFocus: DoEvents
'                                 Call PB_ENT_Click
'                                 Dummy.Enabled = False
                                  Call PB_ENT_Click
'Å¢Å¢---------------------------------------------<< 2005/03/22 >> Update End
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
    Dummy.Enabled = True: Dummy.SetFocus
    
    If wShift = 0 Then
        gSL_Select = "Select ì‡é¶îNåé, åvâÊîNåé, ïiî‘, éËîzêÊ, éÛì¸, çHãÊ From ê∂éYåvâÊÉeÅ[ÉuÉã"
        gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        gSL_Select = gSL_Select & " And   åvâÊîNåé = ì‡é¶îNåé"
        gSL_Select = gSL_Select & " And   SubString( ïiî‘   + Space(25), 1, 25 )"
        gSL_Select = gSL_Select & "     + SubString( éËîzêÊ + Space(7),  1,  7 )"
        gSL_Select = gSL_Select & "     + SubString( éÛì¸   + Space(2),  1,  2 )"
        gSL_Select = gSL_Select & "       < '"
        gSL_Select = gSL_Select & RTrim$(iH1_Hinbn) & Space(25 - Len(RTrim$(iH1_Hinbn)))
        gSL_Select = gSL_Select & RTrim$(iH1_Torcd) & Space(7 - Len(RTrim$(iH1_Torcd)))
        gSL_Select = gSL_Select & RTrim$(iH1_Ukeir) & Space(2 - Len(RTrim$(iH1_Ukeir))) & "'"
        gSL_Select = gSL_Select & " ORDER BY ïiî‘ Desc, éËîzêÊ Desc, éÛì¸ Desc"
    Else
        gSL_Select = "Select ì‡é¶îNåé, åvâÊîNåé, ïiî‘, éËîzêÊ, éÛì¸, çHãÊ From ê∂éYåvâÊÉeÅ[ÉuÉã"
        gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim$(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim$(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim$(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & " And   åvâÊîNåé < '" & RTrim$(iH1_YM) & "'"
        gSL_Select = gSL_Select & " ORDER BY åvâÊîNåé Desc"
    End If
    
    Call DBInput("Back")
    iH1_YM.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    ëO  êi  ÉL Å[(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    Dummy.Enabled = True: Dummy.SetFocus
    
    If wShift = 0 Then
        gSL_Select = "Select ì‡é¶îNåé, åvâÊîNåé, ïiî‘, éËîzêÊ, éÛì¸, çHãÊ From ê∂éYåvâÊÉeÅ[ÉuÉã"
        gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        gSL_Select = gSL_Select & " And   åvâÊîNåé = ì‡é¶îNåé"
        gSL_Select = gSL_Select & " And   SubString( ïiî‘   + Space(25), 1, 25 )"
        gSL_Select = gSL_Select & "     + SubString( éËîzêÊ + Space(7),  1,  7 )"
        gSL_Select = gSL_Select & "     + SubString( éÛì¸   + Space(2),  1,  2 )"
        gSL_Select = gSL_Select & "       > '"
        gSL_Select = gSL_Select & RTrim$(iH1_Hinbn) & Space(25 - Len(RTrim$(iH1_Hinbn)))
        gSL_Select = gSL_Select & RTrim$(iH1_Torcd) & Space(7 - Len(RTrim$(iH1_Torcd)))
        gSL_Select = gSL_Select & RTrim$(iH1_Ukeir) & Space(2 - Len(RTrim$(iH1_Ukeir))) & "'"
        gSL_Select = gSL_Select & " ORDER BY ïiî‘, éËîzêÊ, éÛì¸"
    Else
        gSL_Select = "Select ì‡é¶îNåé, åvâÊîNåé, ïiî‘, éËîzêÊ, éÛì¸, çHãÊ From ê∂éYåvâÊÉeÅ[ÉuÉã"
        gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim$(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim$(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim$(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & " And   åvâÊîNåé > '" & RTrim$(iH1_YM) & "'"
        gSL_Select = gSL_Select & " ORDER BY åvâÊîNåé"
    End If
    
    Call DBInput("Next")
    iH1_YM.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    åü  çı  ÉL Å[(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_SKTable: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_SKTable: Exit Sub
    If PB_Look.Tag = "iH1_KOKbn" Then Call Look_SKTable: Exit Sub
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
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RV_Ukeir
        sv_Hinbn = RV_Hinbn
    End If
End Sub
'+-----------------------------------+
'+       éËîzêÊåüçı
'+-----------------------------------+
Private Function Look_SKTable() As Boolean
    Look_SKTable = False
    RV_Left = 0
    RV_Top = 0
    RV_NJYM = RTrim(iH1_NJYM)
    RV_YM = RTrim(iH1_YM)
    RV_Hinbn = RTrim(iH1_Hinbn)
    If SyoriKB = "ADD" Then RV_Kubun = "MST" Else RV_Kubun = ""
    
    CNJ0025.Show vbModal
    Unload CNJ0025
    Set CNJ0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_Torcd
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RV_Ukeir
        Look_SKTable = True
    End If
End Function
'+----------------------------+
'+    ïœçXëO (F5)
'+----------------------------+
Private Sub PB_REP_Click()
    wInt = 0
    On Error Resume Next: wInt = ActiveControl.Index: On Error GoTo 0
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
        
    For gInt = 0 To 30
        iB1_SKSu(gInt) = CisFun.RSetFld(iB1_SKSu(gInt).Tag, 10, "#,###")
    Next gInt
    B1lb_GSKSu(0) = CisFun.RSetFld(B1lb_GSKSu(0).Tag, 11, "#,###")
    
    If wInt = 0 Then wInt = 1
    iB1_SKSu(wInt).SetFocus
    
    Dummy.Enabled = False
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

    H1lb_Torcd = ""
    If TorNmGet(iH1_Torcd, 9) Then
        H1lb_Torcd = RTrim(TRM_RName)
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
'+  iH1_KOKbn
'+-----------------------------+
Private Sub iH1_KOKbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_KOKbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Sur
'+-----------------------------+
Private Sub iB1_SKSu_LostFocus(Index As Integer)
    gLong = 0
    For gInt = 0 To 30
        gLong = gLong + CisFun.Val2(iB1_SKSu(gInt))
    Next gInt
    B1lb_GSKSu(0) = CisFun.RSetFld(gLong, 11, "#,###")
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
' *CHG* 2004/04/07
'    PB_CAN.Visible = Not H1Mode
    If cmd_KB = 0 Then
        PB_CAN.Visible = Not H1Mode
    Else
        PB_CAN.Visible = False
    End If
    
    PB_REP.Visible = B1Mode
    PB_Look.Visible = False
    
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
    If HeadBody = "B" Then
       For gLong = 0 To 30
           B1lb_ZSKSu(gLong).BackColor = mNInitColor
           B1lb_ZNJSu(gLong).BackColor = mNInitColor
           B1lb_ZSKSu(gLong).Tag = ""
           iB1_SKSu(gLong).Tag = ""
       Next gLong
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
'****************************
'*      HEAD1¡™Ø∏èàóù
'****************************
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

' * éËîzêÊ é©ìÆï\é¶
    If sv_Hinbn <> RTrim$(iH1_Hinbn) Then
        If Not SetTorcd Then Exit Function
    End If
'' * éËîzêÊ
'    If Trim(iH1_Torcd) <> "" Then
'        gSL_Select = " SELECT * "
'        gSL_Select = gSL_Select & " FROM éÊà¯êÊÉ}ÉXÉ^ "
'        gSL_Select = gSL_Select & " WHERE éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & "'"
'        gSL_Select = gSL_Select & "   And éÊà¯êÊãÊï™ = 1"
'        If Not TRMRead(gSL_Select, 1) Then
'            With CisFun
'                .MB_Lines = 4
'                .MB_MSG(2) = "    éËîzêÊ(éÊà¯êÊÉ}ÉXÉ^ñ¢ìoò^)            "
'                .MB_Title = ""
'                .MB_Button = Error
'                .MBOX
'            End With
'           iH1_Torcd.SetFocus
'           Exit Function
'        End If
'    End If
'' * éÛì¸
'    If Trim(iH1_Ukeir) <> "" Then
'        gSL_Select = " SELECT * "
'        gSL_Select = gSL_Select & " FROM éÊà¯êÊéÛì¸É}ÉXÉ^ "
'        gSL_Select = gSL_Select & " Where éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & "'"
'        gSL_Select = gSL_Select & " And   éÊà¯êÊãÊï™ = 1"
'        gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
'        If Not TUMRead(gSL_Select, 1) Then
'            With CisFun
'                .MB_Lines = 4
'                .MB_MSG(2) = "    éÛì¸(éÊà¯êÊéÛì¸É}ÉXÉ^ñ¢ìoò^)            "
'                .MB_Title = ""
'                .MB_Button = Error
'                .MBOX
'            End With
'           iH1_Ukeir.SetFocus
'           Exit Function
'        End If
'    End If
    
'
    gSL_Select = "Select * From ê∂éYåvâÊÉeÅ[ÉuÉã"
    gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    If DBInput("INV") Then
       If SyoriKB = "ADD" Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    ê∂éYåvâÊÉfÅ[É^ ìoò^çœ            "
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
              .MB_MSG(2) = "    ê∂éYåvâÊÉfÅ[É^ ñ¢ìoò^            "
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
'*      HEAD1¡™Ø∏èàóù
'****************************
Private Function Head1Chk_Cmd() As Boolean
    Head1Chk_Cmd = False
    
' * ì‡é¶îNåé
    If Trim(iH1_NJYM) = "" Then GoTo Head1Chk_Cmd_Err
    
' * ëŒè€îNåé
    If Trim(iH1_YM) = "" Then GoTo Head1Chk_Cmd_Err
    
' * ïiî‘
    If Trim(iH1_Hinbn) = "" Then GoTo Head1Chk_Cmd_Err
    
'' * éËîzêÊ
'    If Trim(iH1_Torcd) <> "" Then
'        gSL_Select = " SELECT * "
'        gSL_Select = gSL_Select & " FROM éÊà¯êÊÉ}ÉXÉ^ "
'        gSL_Select = gSL_Select & " WHERE éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & "'"
'        gSL_Select = gSL_Select & "   And éÊà¯êÊãÊï™ = 1"
'        If Not TRMRead(gSL_Select, 1) Then GoTo Head1Chk_Cmd_Err
'    End If
'' * éÛì¸
'    If Trim(iH1_Ukeir) <> "" Then
'        gSL_Select = " SELECT * "
'        gSL_Select = gSL_Select & " FROM éÊà¯êÊéÛì¸É}ÉXÉ^ "
'        gSL_Select = gSL_Select & " Where éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & "'"
'        gSL_Select = gSL_Select & " And   éÊà¯êÊãÊï™ = 1"
'        gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
'        If Not TUMRead(gSL_Select, 1) Then GoTo Head1Chk_Cmd_Err
'    End If
    
'
    gSL_Select = "Select * From ê∂éYåvâÊÉeÅ[ÉuÉã"
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
    Call HeadBodyClear("BX")
    
    If Not SKKRead(gSL_Select, 1) Then
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
        iH1_NJYM = RTrim(SKK.ì‡é¶îNåé)
        iH1_YM = RTrim(SKK.åvâÊîNåé)
        iH1_Hinbn = RTrim(SKK.ïiî‘)
        sv_Hinbn = RTrim(SKK.ïiî‘)
        iH1_Torcd = RTrim(SKK.éËîzêÊ)
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RTrim(SKK.éÛì¸)
        H1lb_KOKbn = RTrim(SKK.çHãÊ)

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
'****************************************************
'*          BODY ì‡óeæØƒèàóù  ( ÉJÉåÉìÉ_ïî )         *
'****************************************************
Private Sub BodySet()
    Static wsNJYM       As String
    Static wsYM       As String
    
' *** ì˙ó êî ***
' * ç°âÒê∂éYåvâÊ
    gSL_Select = "Select * From ê∂éYåvâÊÉeÅ[ÉuÉã"
    gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    If SKKRead(gSL_Select) Then
        Do Until Not SKK_RDSTS
            gInt = CisFun.Val2(Mid(SKK.î[ä˙, 7, 2)) - 1
            iB1_SKSu(gInt) = CisFun.RSetFld(SKK.êîó , 10, "#,###")
            iB1_SKSu(gInt).Tag = SKK.êîó 
            Call SKKReadNext
        Loop
    End If
    Call SKKClose

' * ëOâÒê∂éYåvâÊ
    gSL_Select = "Select * From ê∂éYåvâÊëOâÒÉeÅ[ÉuÉã"
    gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    If SKZRead(gSL_Select) Then
        Do Until Not SKZ_RDSTS
            gInt = CisFun.Val2(Mid(SKZ.î[ä˙, 7, 2)) - 1
            B1lb_ZSKSu(gInt) = CisFun.RSetFld(SKZ.êîó , 10, "#,###")
            B1lb_ZSKSu(gInt).Tag = "X"
            If CisFun.Val2(iB1_SKSu(gInt)) <> SKZ.êîó  Then
                B1lb_ZSKSu(gInt).BackColor = mSHenColor
            End If
            Call SKZReadNext
        Loop
    End If
    Call SKZClose
    
' * ç°âÒì‡é¶(ê∂éYåvâÊïKóvêî)
    gSL_Select = "Select * From ê∂éYåvâÊïKóvêîÉeÅ[ÉuÉã"
    gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    If SHHRead(gSL_Select) Then
        Do Until Not SHH_RDSTS
            gInt = CisFun.Val2(Mid(SHH.î[ä˙, 7, 2)) - 1
            B1lb_NJSu(gInt) = CisFun.RSetFld(SHH.êîó , 10, "#,###")
            
            Call SHHReadNext
        Loop
    End If
    Call SHHClose

' * ëOâÒì‡é¶(ê∂éYåvâÊïKóvêî)
    gSL_Select = "Select * From ê∂éYåvâÊïKóvêîëOâÒÉeÅ[ÉuÉã"
    gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
    gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    If SHZRead(gSL_Select) Then
        Do Until Not SHZ_RDSTS
            gInt = CisFun.Val2(Mid(SHZ.î[ä˙, 7, 2)) - 1
            B1lb_ZNJSu(gInt) = CisFun.RSetFld(SHZ.êîó , 10, "#,###")
            
            Call SHZReadNext
        Loop
    End If
    Call SHZClose
    
' *** åé ó  êî ***
    For gInt = 0 To 2
        If gInt = 0 Then
            gStr = iH1_YM
        Else
            With CisFun
                .Date_RsFormat = fYM
                .Date_Add iH1_YM, M, gInt
                gStr = .Date_Result
            End With
        End If
    ' * ç°âÒåvâÊ
        gSL_Select = "Select SUM(êîó ) êîó  From ê∂éYåvâÊÉeÅ[ÉuÉã"
        gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(gStr) & "'"
        gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & " GROUP BY ì‡é¶îNåé,åvâÊîNåé,ïiî‘,éËîzêÊ,éÛì¸"
        If SKKRead(gSL_Select) Then
            B1lb_GSKSu(gInt) = CisFun.RSetFld(SKK.êîó , 11, "#,###")
            B1lb_GSKSu(gInt).Tag = SKK.êîó 
        End If
    ' * ëOâÒåvâÊ
        gSL_Select = "Select SUM(êîó ) êîó  From ê∂éYåvâÊëOâÒÉeÅ[ÉuÉã"
        gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(gStr) & "'"
        gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & " GROUP BY ì‡é¶îNåé,åvâÊîNåé,ïiî‘,éËîzêÊ,éÛì¸"
        If SKZRead(gSL_Select) Then
            B1lb_GZSKSu(gInt) = CisFun.RSetFld(SKZ.êîó , 11, "#,###")
            B1lb_GZSKSu(gInt).Tag = SKZ.êîó 
        End If
    ' * ç°âÒì‡é¶
        gSL_Select = "Select SUM(êîó ) êîó  From ê∂éYåvâÊïKóvêîÉeÅ[ÉuÉã"
        gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(gStr) & "'"
        gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & " GROUP BY ì‡é¶îNåé,åvâÊîNåé,ïiî‘,éËîzêÊ,éÛì¸"
        If SHHRead(gSL_Select) Then
            B1lb_GNJSu(gInt) = CisFun.RSetFld(SHH.êîó , 11, "#,###")
            B1lb_GNJSu(gInt).Tag = SHH.êîó 
        End If
    ' * ëOâÒì‡é¶
        gSL_Select = "Select SUM(êîó ) êîó  From ê∂éYåvâÊïKóvêîëOâÒÉeÅ[ÉuÉã"
        gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        gSL_Select = gSL_Select & " And   åvâÊîNåé = '" & RTrim(gStr) & "'"
        gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & " GROUP BY ì‡é¶îNåé,åvâÊîNåé,ïiî‘,éËîzêÊ,éÛì¸"
        If SHZRead(gSL_Select) Then
            B1lb_GZNJSu(gInt) = CisFun.RSetFld(SHZ.êîó , 11, "#,###")
            B1lb_GZNJSu(gInt).Tag = SHZ.êîó 
        End If
    Next gInt

' * ïiî‘èÓïÒ
    gSL_Select = "SELECT HN.*, "
    gSL_Select = gSL_Select & " ISNULL(SY.ílñºèÃ,'') ç›å…ä«óùãÊï™ñº "
    gSL_Select = gSL_Select & " FROM ïiî‘É}ÉXÉ^ HN "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN ñºèÃÉ}ÉXÉ^ SY "
    gSL_Select = gSL_Select & "   ON SY.ãÊï™ñºèÃ = 'ç›å…ãÊï™' "
    gSL_Select = gSL_Select & "  AND SY.íl = HN.ç›å…ä«óùãÊï™ "
    gSL_Select = gSL_Select & "  AND SY.ãÊï™É^ÉCÉv = '' "
    gSL_Select = gSL_Select & " WHERE HN.ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    If HNMRead(gSL_Select, 1) Then
        B2lb_ZKKbn = HNM.ç›å…ä«óùãÊï™ñº
        Select Case HNM.ç›å…ä«óùãÊï™
            Case 2: B2lb_ZKSu = Format(HNM.ç›å…ä«óùêîíl, "###") & " Åì"
            Case 3: B2lb_ZKSu = Format(HNM.ç›å…ä«óùêîíl, "###,###")
            Case 4: B2lb_ZKSu = Format(HNM.ç›å…ä«óùêîíl, "#0.0") & " ì˙"
            Case Else: B2lb_ZKSu = ""
        End Select
    End If
End Sub
'****************************************************
'*          BODY ì‡óeæØƒèàóù  ( ïiî‘èÓïÒïî )         *
'****************************************************
Private Sub Body12Set()
    Static wsNJYM       As String
    Static wsYM       As String
' * ïiî‘éÊà¯êÊ
    gSL_Select = "SELECT * "
    gSL_Select = gSL_Select & " FROM ïiî‘éÊà¯êÊÉ}ÉXÉ^  "
    gSL_Select = gSL_Select & " WHERE ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   AND éÊà¯êÊ = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   AND ïiñ⁄ = 1 "
    If HTMRead(gSL_Select, 1) Then
        
    ' ÉäÅ[ÉhÉ^ÉCÉÄ
        B1lb_LTime = CisFun.RSetFld(HTM.LT, 2, "##")
    ' ê∂éYÉçÉbÉg
        B1lb_Lot = CisFun.RSetFld(HTM.ÉçÉbÉg1, 6, "#,###")
    ' ójì˙
        If HTM.Week1 <> 0 Then: B1lb_Week(1) = "åé"
        If HTM.Week2 <> 0 Then: B1lb_Week(2) = "âŒ"
        If HTM.Week3 <> 0 Then: B1lb_Week(3) = "êÖ"
        If HTM.Week4 <> 0 Then: B1lb_Week(4) = "ñÿ"
        If HTM.Week5 <> 0 Then: B1lb_Week(5) = "ã‡"
    ' * ì˙ït
        If HTM.Day1 <> 0 Then: B1lb_Day(1) = CisFun.RSetFld(HTM.Day1, 2, "##")
        If HTM.Day2 <> 0 Then: B1lb_Day(2) = CisFun.RSetFld(HTM.Day2, 2, "##")
        If HTM.Day3 <> 0 Then: B1lb_Day(3) = CisFun.RSetFld(HTM.Day3, 2, "##")
        If HTM.Day4 <> 0 Then: B1lb_Day(4) = CisFun.RSetFld(HTM.Day4, 2, "##")
        If HTM.Day5 <> 0 Then: B1lb_Day(5) = CisFun.RSetFld(HTM.Day5, 2, "##")
    ' * SKãÊï™
        If HTM.SKãÊï™ = 0 Then
            B2lb_SKKbn = "Ç©ÇÒÇŒÇÒ"
        Else
            B2lb_SKKbn = "éw é¶"
        End If
    ' * ÉTÉCÉNÉã
        If RTrim(HTM.ÉTÉCÉNÉã) <> "" Then
            B2lb_Cycle = Mid(HTM.ÉTÉCÉNÉã, 1, 1) & "-" & Mid(HTM.ÉTÉCÉNÉã, 2, 2) & "-" & Mid(HTM.ÉTÉCÉNÉã, 4, 2)
        End If
    ' * é˚óeêî1
        If HTM.é˚óeêî1 <> 0 Then: B2lb_Syuyo1 = CisFun.RSetFld(HTM.é˚óeêî1, 7, "#,###")
    ' * êÿë÷ì˙ÅEï÷
        If RTrim(HTM.ïœçXì˙) <> "" Then
            B2lb_CDate = Mid(HTM.ïœçXì˙, 1, 4) & "/" & Mid(HTM.ïœçXì˙, 5, 2) & "/" & Mid(HTM.ïœçXì˙, 7, 2)
        End If
        If HTM.ïœçXï÷ <> "" Then: B2lb_CBin = CisFun.RSetFld(HTM.ïœçXï÷, 2, "##")
    ' * é˚óeêî2
        If HTM.é˚óeêî2 <> 0 Then: B2lb_Syuyo2 = CisFun.RSetFld(HTM.é˚óeêî2, 7, "#,###")
    ' * åWêî1
        If HTM.åWêî1 <> 0 Then: B2lb_Kesu1 = CisFun.RSetFld(HTM.åWêî1, 4, "0.00")
    ' * åWêî2
        If HTM.åWêî2 <> 0 Then: B2lb_Kesu2 = CisFun.RSetFld(HTM.åWêî2, 4, "0.00")
    ' * åWêî3
        If HTM.åWêî3 <> 0 Then: B2lb_Kesu3 = CisFun.RSetFld(HTM.åWêî3, 4, "0.00")
    ' * î‰ó¶
        If HTM.î≠íçî‰ó¶ <> 0 Then: B2lb_Ritu = CisFun.RSetFld(HTM.î≠íçî‰ó¶, 3, "#") & "%"
    ' * óLå¯ä˙å¿
        If RTrim(HTM.óLå¯ä˙å¿äJén) <> "" Then
            B2lb_SDate = Mid(HTM.óLå¯ä˙å¿äJén, 1, 4) & "/" & Mid(HTM.óLå¯ä˙å¿äJén, 5, 2) & "/" & Mid(HTM.óLå¯ä˙å¿äJén, 7, 2)
        End If
        
        If RTrim(HTM.óLå¯ä˙å¿èIóπ) <> "" Then
            B2lb_SDate = Mid(HTM.óLå¯ä˙å¿èIóπ, 1, 4) & "/" & Mid(HTM.óLå¯ä˙å¿èIóπ, 5, 2) & "/" & Mid(HTM.óLå¯ä˙å¿èIóπ, 7, 2)
        End If
        
    End If
    
' * ìñåéì‡é¶êî
    gSL_Select = "SELECT SUM(êîó ) êîó "
    gSL_Select = gSL_Select & " FROM ê∂éYåvâÊïKóvêîÉeÅ[ÉuÉã "
    gSL_Select = gSL_Select & " WHERE ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
    gSL_Select = gSL_Select & "   AND åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "   AND ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " GROUP BY ì‡é¶îNåé,åvâÊîNåé,ïiî‘ "
    If SHHRead(gSL_Select, 1) Then
        B2lb_GNJSu = CisFun.RSetFld(SHH.êîó , 11, "#,###")
    End If
' * ëOåéì‡é¶êî
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add iH1_YM, M, -1
        wsYM = .Date_Result
    End With
    If Trim(iH1_NJYM) = Trim(iH1_YM) Then
        wsNJYM = wsYM
    Else
        wsNJYM = iH1_NJYM
    End If
    gSL_Select = "SELECT SUM(êîó ) êîó "
    gSL_Select = gSL_Select & " FROM ê∂éYåvâÊïKóvêîÉeÅ[ÉuÉã "
    gSL_Select = gSL_Select & " WHERE ì‡é¶îNåé = '" & RTrim(wsNJYM) & "'"
    gSL_Select = gSL_Select & "   AND åvâÊîNåé = '" & RTrim(wsYM) & "'"
    gSL_Select = gSL_Select & "   AND ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " GROUP BY ì‡é¶îNåé,åvâÊîNåé,ïiî‘ "
    If SHHRead(gSL_Select, 1) Then
        B2lb_GZNJSu = CisFun.RSetFld(SHH.êîó , 11, "#,###")
    End If
' * ì‡é¶ç∑
    B2lb_GSNJSu = CisFun.RSetFld(CisFun.Val2(B2lb_GNJSu) - CisFun.Val2(B2lb_GZNJSu), 11, "#,###")

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
    
        .ParaNo = 5:    B2lb_Zaiko = Format(.ParaValue, "#,###")
    End With
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
        .SQL = "DELETE FROM ê∂éYåvâÊÉeÅ[ÉuÉã "
        .SQL = .SQL & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        .SQL = .SQL & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
        .SQL = .SQL & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        .SQL = .SQL & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
        .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
        .DBExec
        
        .SQL = "DELETE FROM ê∂éYåvâÊëOâÒÉeÅ[ÉuÉã "
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
' *ADD* 2004/04/09
    Dim MntFlg          As Boolean

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

' *ADD* 2004/04/09
    MntFlg = False
    For gInt = 0 To 30
        If CisFun.Val2(iB1_SKSu(gInt)) <> CisFun.Val2(iB1_SKSu(gInt).Tag) Then
            MntFlg = True
            Exit For
        End If
    Next gInt

    If MntFlg And SyoriKB = "MNT" Then
    ' *ëOâÒì‡óe çÌèú*
        With CisDB
            .SQL = "DELETE FROM ê∂éYåvâÊëOâÒÉeÅ[ÉuÉã "
            .SQL = .SQL & " Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
            .SQL = .SQL & " And   åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
            .SQL = .SQL & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
            .SQL = .SQL & " And   éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
            .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
            .DBExec
        End With
        
    ' *ëOâÒÉfÅ[É^ çÏê¨*
        With CisDB
            .SQL = "INSERT INTO ê∂éYåvâÊëOâÒÉeÅ[ÉuÉã "
            .SQL = .SQL & " (ì‡é¶îNåé, ïiî‘, éËîzêÊ, éÛì¸, çHãÊ, "
            .SQL = .SQL & "  î[ä˙, édä|ì˙, éËîzì˙, êîó , "
            .SQL = .SQL & "  èàóùãÊï™, çÌèúãÊï™, ïœçXâÒêî, åvâÊîNåé, "
            .SQL = .SQL & "  çÏê¨é“, çÏê¨ì˙, çÏê¨í[ññ )"
            .SQL = .SQL & "  SELECT ì‡é¶îNåé, ïiî‘, éËîzêÊ, éÛì¸, çHãÊ, "
            .SQL = .SQL & "         î[ä˙, édä|ì˙, éËîzì˙, êîó , "
            .SQL = .SQL & "         èàóùãÊï™, çÌèúãÊï™, ïœçXâÒêî, åvâÊîNåé, "
            .SQL = .SQL & "'" & gTanto & "',"
            .SQL = .SQL & "GETDATE(),HOST_NAME() "
            .SQL = .SQL & "         FROM ê∂éYåvâÊÉeÅ[ÉuÉã "
            .SQL = .SQL & "         Where ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
            .SQL = .SQL & "           And åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
            .SQL = .SQL & "           And ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
            .SQL = .SQL & "           And éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
            .SQL = .SQL & "           And éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
            .DBExec
        End With
    End If

'   çÄñ⁄à⁄ëó
    ClwFlg = False
    For gInt = 0 To 30
        If CisFun.Val2(iB1_SKSu(gInt)) <> CisFun.Val2(iB1_SKSu(gInt).Tag) Then
            
            gStr = RTrim(iH1_YM) & Format(gInt + 1, "00")
            gSL_Select = "SELECT * FROM ê∂éYåvâÊÉeÅ[ÉuÉã "
            gSL_Select = gSL_Select & " WHERE ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
            gSL_Select = gSL_Select & "   AND ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
            gSL_Select = gSL_Select & "   AND éËîzêÊ = '" & RTrim(iH1_Torcd) & "'"
            gSL_Select = gSL_Select & "   AND éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
            gSL_Select = gSL_Select & "   AND î[ä˙ = '" & RTrim(gStr) & "'"
            If SKKRead(gSL_Select, 1) Then
                AddFlg = False
' *DEL* 2004/04/09
'                With SKK
'                    SKZ.ì‡é¶îNåé = .ì‡é¶îNåé
'                    SKZ.ïiî‘ = .ïiî‘
'                    SKZ.éËîzêÊ = .éËîzêÊ
'                    SKZ.éÛì¸ = .éÛì¸
'                    SKZ.çHãÊ = .çHãÊ
'                    SKZ.î[ä˙ = .î[ä˙
'                    SKZ.édä|ì˙ = .édä|ì˙
'                    SKZ.éËîzì˙ = .éËîzì˙
'                    SKZ.êîó  = .êîó 
'                    SKZ.èàóùãÊï™ = .èàóùãÊï™
'                    SKZ.çÌèúãÊï™ = .çÌèúãÊï™
'                    SKZ.ïœçXâÒêî = 1
'                    SKZ.åvâÊîNåé = .åvâÊîNåé
'                End With
'
'                If RTrim(B1lb_ZSKSu(gInt).Tag) = "" Then
'                    SKZ.çÏê¨é“ = gTanto
'                    Call SKZInsert
'                Else
'                    SKZ.çXêVé“ = gTanto
'                    Call SKZUpdate
'                End If
            Else
                Call ItemsClearSKK
                
            ' ÉÅÉìÉeópÉJÉåÉìÉ_ÉèÅ[ÉNçÏê¨
                If Not ClwFlg Then
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
                    ClwFlg = True
                End If
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
                With SKK
                    .ì‡é¶îNåé = RTrim(iH1_NJYM)
                    .ïiî‘ = RTrim(iH1_Hinbn)
                    .éËîzêÊ = RTrim(iH1_Torcd)
                    .éÛì¸ = RTrim(iH1_Ukeir)
'                    .çHãÊ = RTrim(iH1_KOKbn)
                    .î[ä˙ = RTrim(gStr)
                    .édä|ì˙ = RTrim(wsSYMD)
                    .éËîzì˙ = RTrim(wsTYMD)
                End With
            End If
    
            SKK.êîó  = CisFun.Val2(iB1_SKSu(gInt))
            SKK.åvâÊîNåé = RTrim(iH1_YM)
            If SKK.êîó  <> 0 Then
                If AddFlg Then
                    SKK.çÏê¨é“ = gTanto
                    Call SKKInsert
                    iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
                Else
                    SKK.çXêVé“ = gTanto
                    Call SKKUpdate
                    iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
                End If
            Else
            
        ' ( çÌèúèàóù )
                With CisDB
                    .SQL = "DELETE FROM ê∂éYåvâÊÉeÅ[ÉuÉã "
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
        End If
    Next gInt
    
'
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
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
    For gInt = 0 To 30
        If gInt <= CLM.ññì˙ - 1 Then
            B_Sur(gInt).Visible = True
            
            Select Case CisFun.Val2(Mid(CLM.â“ìÆãÊï™, gInt + 1, 1)) _
                      + CisFun.Val2(Mid(CLM.ñÈãŒãÊï™, gInt + 1, 1))
            
                Case 0
'                       < ëSâ“ìÆ >
                        B1lb_Kad(gInt).Visible = False
                        BLB_Day(gInt).BackColor = mOndayColor
                Case 1
'                       < îºâ“ìÆ >
                        B1lb_Kad(gInt).Visible = True
                        BLB_Day(gInt).BackColor = mOndayColor
                        If Mid(CLM.â“ìÆãÊï™, gInt + 1, 1) = "1" Then B1lb_Kad(gInt) = "ñÈ" Else B1lb_Kad(gInt) = "íã"
                Case Else
'                       < ãxì˙ >
                        B1lb_Kad(gInt).Visible = False
                        BLB_Day(gInt).BackColor = mOffdayColor
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
'                sv_Hinbn = RTrim$(iH1_Hinbn)
                If Not Look_SKTable Then Exit Function
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
        gSL_Select = gSL_Select & " FROM ê∂éYåvâÊÉeÅ[ÉuÉã "
        gSL_Select = gSL_Select & " WHERE ì‡é¶îNåé = '" & RTrim(iH1_NJYM) & "'"
        gSL_Select = gSL_Select & "   AND åvâÊîNåé = '" & RTrim(iH1_YM) & "'"
        gSL_Select = gSL_Select & "   AND ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " GROUP BY éËîzêÊ, éÛì¸"
        If SKKRead(gSL_Select) Then
        ' * éËîzêÊëIë
            If CisDB.RecordCount > 1 Then
                Call SKKClose
'                sv_Hinbn = RTrim$(iH1_Hinbn)
                If Not Look_SKTable Then Exit Function
                sv_Hinbn = RTrim$(iH1_Hinbn)
            Else
                iH1_Torcd = SKK.éËîzêÊ
                iH1_Ukeir = SKK.éÛì¸
                Call SKKClose
            End If
        End If
    End If

SetTorcd_ED:
    SetTorcd = True
End Function
    

