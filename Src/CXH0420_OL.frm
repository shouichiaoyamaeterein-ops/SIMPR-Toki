VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CisYMDwB_6.0.ocx"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CXH0420 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "î≠íçåvâÊèÛãµè∆âÔ"
   ClientHeight    =   11190
   ClientLeft      =   60
   ClientTop       =   330
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
   WindowState     =   2  'ç≈ëÂâª
   Begin VB.PictureBox B3_Area2 
      Height          =   3750
      Left            =   6090
      ScaleHeight     =   3690
      ScaleWidth      =   4380
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5670
      Width           =   4440
      Begin Cis3D_v60.CIS3D CIS3D42 
         Height          =   795
         Left            =   2610
         Top             =   0
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1402
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
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         Begin Cis3D_v60.CIS3D B3lb_LTime 
            Height          =   345
            Left            =   645
            Top             =   375
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
      End
      Begin Cis3D_v60.CIS3D Back_Week 
         Height          =   630
         Left            =   15
         Top             =   3045
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1111
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
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D B3lb_Week 
            Height          =   345
            Index           =   1
            Left            =   1245
            Top             =   135
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
         Begin Cis3D_v60.CIS3D B3lb_Week 
            Height          =   345
            Index           =   2
            Left            =   1695
            Top             =   135
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
         Begin Cis3D_v60.CIS3D B3lb_Week 
            Height          =   345
            Index           =   3
            Left            =   2145
            Top             =   135
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
         Begin Cis3D_v60.CIS3D B3lb_Week 
            Height          =   345
            Index           =   4
            Left            =   2595
            Top             =   135
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
         Begin Cis3D_v60.CIS3D B3lb_Week 
            Height          =   345
            Index           =   5
            Left            =   3045
            Top             =   135
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
      Begin Cis3D_v60.CIS3D CIS3D29 
         Height          =   795
         Left            =   1350
         Top             =   0
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
         Begin Cis3D_v60.CIS3D B3lb_Cycle 
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
         Left            =   15
         Top             =   795
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
         Begin Cis3D_v60.CIS3D B3lb_Syuyo1 
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
      Begin Cis3D_v60.CIS3D CIS3D10 
         Height          =   1590
         Left            =   1245
         Top             =   795
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   2805
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
         cPositionY      =   210
         Begin Cis3D_v60.CIS3D B3lb_CDate 
            Height          =   330
            Left            =   90
            Top             =   660
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
         Begin Cis3D_v60.CIS3D B3lb_CBin 
            Height          =   330
            Left            =   1380
            Top             =   660
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
      Begin Cis3D_v60.CIS3D CIS3D15 
         Height          =   660
         Left            =   15
         Top             =   2385
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1164
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
         cAlingnment     =   1
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D B3lb_ZKKbn 
            Height          =   330
            Left            =   1245
            Top             =   180
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
         Begin Cis3D_v60.CIS3D B3lb_ZKSu 
            Height          =   330
            Left            =   2430
            Top             =   180
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
      Begin Cis3D_v60.CIS3D CIS3D16 
         Height          =   795
         Left            =   3135
         Top             =   795
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
         Begin Cis3D_v60.CIS3D B3lb_Syuyo2 
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
      Begin Cis3D_v60.CIS3D CIS3D17 
         Height          =   795
         Left            =   0
         Top             =   0
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
         Begin Cis3D_v60.CIS3D B3lb_SKKbn 
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
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   795
         Left            =   15
         Top             =   1590
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "î≠íç€Øƒá@"
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
         Begin Cis3D_v60.CIS3D B3lb_Lot1 
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
      Begin Cis3D_v60.CIS3D CIS3D25 
         Height          =   795
         Left            =   3135
         Top             =   1590
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "î≠íç€ØƒáA"
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
         Begin Cis3D_v60.CIS3D B3lb_Lot2 
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
   End
   Begin Cis3D_v60.CIS3D PNL_Wait 
      Height          =   930
      Left            =   5340
      Top             =   4380
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
   Begin VB.PictureBox H3_Area1 
      Height          =   3780
      Left            =   105
      ScaleHeight     =   3720
      ScaleWidth      =   4395
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7230
      Width           =   4455
      Begin TrueOleDBGrid60.TDBGrid DBGrid1 
         Bindings        =   "CXH0420.frx":0000
         Height          =   3420
         Left            =   0
         OleObjectBlob   =   "CXH0420.frx":0015
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   300
         Width           =   4380
      End
      Begin Cis3D_v60.CIS3D BLB_D 
         Height          =   315
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   556
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "ç›å…èÓïÒïœçXéñóR"
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
         cFont3D         =   1
         cAlingnment     =   7
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
      Height          =   2580
      Left            =   105
      ScaleHeight     =   2520
      ScaleWidth      =   4425
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1095
      Width           =   4485
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   795
         Left            =   2085
         Top             =   1665
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   1402
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
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   795
         Left            =   3825
         Top             =   870
         Width           =   570
         _ExtentX        =   1005
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
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   375
            Left            =   90
            TabIndex        =   3
            Top             =   345
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
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   795
         Left            =   60
         Top             =   1665
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "éw íË ì˙"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_Ymd 
            Height          =   345
            Left            =   105
            TabIndex        =   4
            Top             =   360
            Width           =   1800
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
            Object.Width           =   1800
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
         Height          =   795
         Left            =   1215
         Top             =   75
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "ïi   î‘"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Left            =   30
            TabIndex        =   1
            Top             =   360
            Width           =   3150
            _ExtentX        =   5556
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   795
         Left            =   60
         Top             =   870
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "éd  ì¸  êÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   60
            TabIndex        =   2
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
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1050
            Top             =   360
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXX"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
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
      Begin Cis3D_v60.CIS3D CIS3D38 
         Height          =   795
         Left            =   60
         Top             =   75
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1402
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
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Seban 
            Height          =   360
            Left            =   45
            TabIndex        =   0
            Top             =   375
            Width           =   1080
            _ExtentX        =   1905
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
            Text            =   "XXXXXXXX"
            MaxLength       =   8
            IMEMode         =   2
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   9930
      Left            =   4635
      ScaleHeight     =   9870
      ScaleWidth      =   10440
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1095
      Width           =   10500
      Begin vsFlexLib.vsFlexArray vsGrid 
         Height          =   9795
         Left            =   75
         TabIndex        =   5
         Top             =   45
         Width           =   10290
         _Version        =   196608
         _ExtentX        =   18150
         _ExtentY        =   17277
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   "ì˙ït    | |ÇPì˙      |ÇQì˙      |ÇRì˙      |ÇSì˙       |ÇTì˙    |ÇUì˙    |ÇVì˙      |"
         Rows            =   36
         Cols            =   11
         BackColor       =   16777152
         BackColorFixed  =   16777152
         BackColorBkg    =   12632256
         TextStyle       =   4
         TextStyleFixed  =   3
         FocusRect       =   2
         HighLight       =   0
         ScrollBars      =   0
         RowHeightMin    =   200
         FillStyle       =   1
         GridColor       =   14737632
         SelectionMode   =   2
         BackColorAlternate=   16777152
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
      Height          =   660
      Left            =   9690
      ScaleHeight     =   600
      ScaleWidth      =   5370
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   390
      Width           =   5430
      Begin CisBtn_60.CisBtn PB_Disp 
         Height          =   585
         Left            =   945
         Top             =   0
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
         PFForeColor     =   65535
         PFShadowColor   =   0
         PFBackColor     =   0
      End
      Begin CisBtn_60.CisBtn PB_Chg 
         Height          =   585
         Left            =   1830
         Top             =   0
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   2715
         Top             =   0
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
         Left            =   3600
         Top             =   0
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
         Left            =   4485
         Top             =   0
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
         Left            =   60
         Top             =   0
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
      TabIndex        =   7
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
   Begin Cis3D_v60.CIS3D UC_3D11 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "Åy î≠íçåvâÊèÛãµè∆âÔ Åz"
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
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14430
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   11910
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   3450
      Left            =   105
      ScaleHeight     =   3390
      ScaleWidth      =   4395
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3750
      Width           =   4455
      Begin Cis3D_v60.CIS3D CIS3D19 
         Height          =   420
         Left            =   2730
         Top             =   1650
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   741
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
         Begin Cis3D_v60.CIS3D B1lb_SSuryo 
            Height          =   315
            Index           =   4
            Left            =   170
            Top             =   45
            Width           =   1300
            _ExtentX        =   2302
            _ExtentY        =   556
            BackColor       =   16777215
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   420
         Left            =   2730
         Top             =   1230
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   741
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
         Begin Cis3D_v60.CIS3D B1lb_SSuryo 
            Height          =   315
            Index           =   3
            Left            =   170
            Top             =   45
            Width           =   1300
            _ExtentX        =   2302
            _ExtentY        =   556
            BackColor       =   12640511
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   420
         Left            =   2730
         Top             =   810
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   741
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
         Begin Cis3D_v60.CIS3D B1lb_SSuryo 
            Height          =   315
            Index           =   2
            Left            =   170
            Top             =   45
            Width           =   1300
            _ExtentX        =   2302
            _ExtentY        =   556
            BackColor       =   12648447
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   420
         Left            =   1125
         Top             =   390
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   741
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
         Begin Cis3D_v60.CIS3D B1lb_TSyoyo 
            Height          =   315
            Index           =   1
            Left            =   170
            Top             =   45
            Width           =   1300
            _ExtentX        =   2302
            _ExtentY        =   556
            BackColor       =   12648384
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D Grid_RowHead 
         Height          =   420
         Index           =   1
         Left            =   45
         Tag             =   "åvâÊêî"
         Top             =   390
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   741
         ForeColor       =   16711680
         Caption         =   "åvâÊêî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      Begin Cis3D_v60.CIS3D Grid_RowHead 
         Height          =   420
         Index           =   2
         Left            =   45
         Tag             =   "î≠íçêî"
         Top             =   810
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   741
         ForeColor       =   16711680
         Caption         =   "î≠íçêî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      Begin Cis3D_v60.CIS3D Grid_RowHead 
         Height          =   420
         Index           =   3
         Left            =   45
         Tag             =   "ì¸â◊êî"
         Top             =   1230
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   741
         ForeColor       =   16711680
         Caption         =   "ì¸â◊êî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      Begin Cis3D_v60.CIS3D Grid_RowHead 
         Height          =   420
         Index           =   4
         Left            =   45
         Tag             =   "é¿ê—êî"
         Top             =   1650
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   741
         ForeColor       =   16711680
         Caption         =   "é¿ê—êî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      Begin Cis3D_v60.CIS3D BLB_D 
         Height          =   2010
         Index           =   17
         Left            =   5415
         Top             =   360
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   3545
         BackColor       =   0
         ForeColor       =   16744703
         Caption         =   "ïiî‘èÓïÒ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   300
      End
      Begin Cis3D_v60.CIS3D BLB_D 
         Height          =   360
         Index           =   18
         Left            =   1110
         Top             =   30
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   635
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "ìñåéçáåv"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Grid_RowHead 
         Height          =   420
         Index           =   5
         Left            =   45
         Tag             =   "ç›å…êî"
         Top             =   2070
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   741
         ForeColor       =   16711680
         Caption         =   "ëOåéññç›å…êî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         cPositionX      =   500
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   420
         Left            =   2730
         Top             =   390
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   741
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
         Begin Cis3D_v60.CIS3D B1lb_SSuryo 
            Height          =   315
            Index           =   1
            Left            =   170
            Top             =   45
            Width           =   1300
            _ExtentX        =   2302
            _ExtentY        =   556
            BackColor       =   12648384
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D BLB_D 
         Height          =   360
         Index           =   0
         Left            =   2730
         Top             =   30
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   635
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "éwíËì˙çáåv"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   1
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   420
         Left            =   2730
         Top             =   2070
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   741
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
         Begin Cis3D_v60.CIS3D B1lb_TSyoyo 
            Height          =   315
            Index           =   5
            Left            =   150
            Top             =   45
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   420
         Left            =   1125
         Top             =   810
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   741
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
         Begin Cis3D_v60.CIS3D B1lb_TSyoyo 
            Height          =   315
            Index           =   2
            Left            =   170
            Top             =   45
            Width           =   1300
            _ExtentX        =   2302
            _ExtentY        =   556
            BackColor       =   12648447
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   420
         Left            =   1125
         Top             =   1230
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   741
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
         Begin Cis3D_v60.CIS3D B1lb_TSyoyo 
            Height          =   315
            Index           =   3
            Left            =   170
            Top             =   45
            Width           =   1300
            _ExtentX        =   2302
            _ExtentY        =   556
            BackColor       =   12640511
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   420
         Left            =   1125
         Top             =   1650
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   741
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
         Begin Cis3D_v60.CIS3D B1lb_TSyoyo 
            Height          =   315
            Index           =   4
            Left            =   170
            Top             =   45
            Width           =   1300
            _ExtentX        =   2302
            _ExtentY        =   556
            BackColor       =   16777215
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D12 
         Height          =   360
         Left            =   45
         Top             =   30
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   635
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
      End
      Begin Cis3D_v60.CIS3D Grid_RowHead 
         Height          =   420
         Index           =   0
         Left            =   45
         Top             =   2490
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   741
         ForeColor       =   16711680
         Caption         =   "óÇåéåvâÊêî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         cPositionX      =   500
      End
      Begin Cis3D_v60.CIS3D CIS3D20 
         Height          =   420
         Left            =   45
         Top             =   2910
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   741
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
         Begin Cis3D_v60.CIS3D B1lb_YSyoyo 
            Height          =   315
            Left            =   450
            Top             =   45
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D Grid_RowHead 
         Height          =   420
         Index           =   6
         Left            =   2205
         Top             =   2490
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   741
         ForeColor       =   16711680
         Caption         =   "óÇÅXåéåvâÊêî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         cPositionX      =   500
      End
      Begin Cis3D_v60.CIS3D CIS3D39 
         Height          =   420
         Left            =   2205
         Top             =   2910
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   741
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
         Begin Cis3D_v60.CIS3D B1lb_YYSyoyo 
            Height          =   315
            Left            =   450
            Top             =   45
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
      End
   End
   Begin Cis3D_v60.CIS3D DispSuryo 
      Height          =   600
      Left            =   105
      Top             =   420
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1058
      BackColor       =   16711680
      ForeColor       =   16777215
      Caption         =   "êî ó "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   21.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      cAlingnment     =   7
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   9375
      Top             =   120
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   635
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Cis3D_v60.CIS3D DispMei 
      Height          =   600
      Left            =   1860
      Top             =   420
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1058
      BackColor       =   12648384
      ForeColor       =   0
      Caption         =   "é¿  ê—"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   21.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      cAlingnment     =   7
   End
   Begin VB.Label ColorToday 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00FFC0C0&
      Caption         =   "ìñì˙êF"
      Height          =   330
      Left            =   8265
      TabIndex        =   16
      Top             =   540
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label ColorSata 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00FF8080&
      Caption         =   "ìyójêF"
      Height          =   330
      Left            =   7095
      TabIndex        =   15
      Top             =   540
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label ColorSunday 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00C0C0FF&
      Caption         =   "ì˙ójêF"
      Height          =   315
      Left            =   5940
      TabIndex        =   14
      Top             =   540
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label DayColor 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00E0E0E0&
      Caption         =   "ëŒè€äOêF"
      Height          =   300
      Left            =   4725
      TabIndex        =   13
      Top             =   540
      Visible         =   0   'False
      Width           =   1050
   End
End
Attribute VB_Name = "CXH0420"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :   î≠íçåvâÊèÛãµè∆âÔ
'**       ÉtÉHÅ[ÉÄID    :   CXH0420
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :   2004/02/18  By CIS
'**       ïœ  çX  ì˙    :
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' Ã´∞—±∏√®Ãﬁ Ã◊∏ﬁ
    Dim ProcHB                  As String * 2   ' ì¸óÕàÊêÿë÷
    
    Dim mWeekStart              As Integer
    Dim mSaturdayColor          As Long
    Dim mSundayColor            As Long

    Dim mDayRow(31)             As Integer
    Dim mDayCol(31)             As Integer

    Dim mSuryo(5, 31)           As Long
    Dim mMaisu(5, 31)           As Long
    Dim mZaiko(31)              As Long
    Dim mNyuka(31)              As Long
    Dim mZaikoNew               As Long
    
    Dim mSuryoK(31)             As Long
    Dim mSuryoA(31)             As Long
    
    Dim mSetType                As Integer
    Dim mCellSet                As Boolean
    Dim mSetDisp                As Integer
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
    
    B3_Area2.Move H3_Area1.Left, H3_Area1.Top
    
    iH1_Ymd = Format(Date, "YYYYMM")
    ' ójì˙äJénÇãÅÇﬂÇÈ
    gSL_Select = "select êîéö1 From ñºèÃÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " Where ãÊï™ñºèÃ = 'ójì˙äJénä«óù'"
    gSL_Select = gSL_Select & "   And íl = '1'"
    If Not SYMRead(gSL_Select, 1) Then
       SYM.êîéö1 = 0
    End If
    
    mWeekStart = SYM.êîéö1 + 1

    mSaturdayColor = &HFFC0C0
    mSundayColor = &HC0C0FF

    Call GridInit
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
    
    mSetType = 1
    Call PB_Chg_Click
    mSetDisp = 2
    Call PB_Disp_Click
' ADO∫›ƒ€∞ŸÅ@ê⁄ë±èÓïÒÅ@ÉZÉbÉg
    Adodc1.ConnectionString = CisDB.ConnectString
    
    FormAct = Not FormAct
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
           Case vbKeyF1:        If PB_Disp.Visible Then Call PB_Disp_Click    'Åyêÿë÷Åz
           Case vbKeyF6:        If PB_Chg.Visible Then Call PB_Chg_Click    'Åyêÿë÷Åz
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  'ÅyåüçıÅz
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
    Call CisDB.DBDISConnect                            ' √ﬁ∞¿Õﬁ∞Ωêÿíf
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
    
    If ProcHB = "H1" Then
        Call HeadBodyClear("H")
        iH1_Ymd = Format(Date, "YYYYMM")
        iH1_Seban.SetFocus
        GoTo PB_CAN_Ed
    End If
    
    If ProcHB = "B1" Then
        ProcHB = "H1"
        PB_Disp.Visible = False
        DispMei.Visible = False
        H3_Area1.Visible = False
        B3_Area2.Visible = True
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        Call GridInit
        Call DBGridClear
        iH1_Seban.SetFocus
        GoTo PB_CAN_Ed
    End If
    
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    êÿÅ@ë÷  ÉL Å[(F6)       +
'+----------------------------+
Private Sub PB_Chg_Click()
    Dim SaveRow As Long
    Dim SaveCol As Long
    
    If mSetType = 0 Then
       mSetType = 1
       DispSuryo = "ñá êî"
       PB_Chg.BtnCaption = "êîó "
    Else
       mSetType = 0
       DispSuryo = "êî ó "
       PB_Chg.BtnCaption = "ñáêî"
    End If
    SaveRow = vsGrid.Row
    SaveCol = vsGrid.Col
    Call GridValueSet
    vsGrid.Row = SaveRow
    vsGrid.Col = SaveCol
    
End Sub
'+----------------------------+
'+    êÿÅ@ë÷  ÉL Å[(F1)       +
'+----------------------------+
Private Sub PB_Disp_Click()
    Dim SaveRow As Long
    Dim SaveCol As Long
    
    Select Case mSetDisp
        Case 0
            mSetDisp = 1
            DispMei = "åv  âÊ"
            PB_Disp.BtnCaption = "é¿•åv"
        Case 1
            mSetDisp = 2
            DispMei = "é¿ê—•åvâÊ"
            PB_Disp.BtnCaption = "é¿ê—"
        Case Else
            mSetDisp = 0
            DispMei = "é¿  ê—"
            PB_Disp.BtnCaption = "åvâÊ"
    End Select
    SaveRow = vsGrid.Row
    SaveCol = vsGrid.Col
    Call GridValueSet
    vsGrid.Row = SaveRow
    vsGrid.Col = SaveCol
    
End Sub
'+----------------------------+
'+    åü  çı  ÉL Å[(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Ymd" Then iH1_Ymd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_SSTorcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_UTorcd(0): Exit Sub
    If PB_Look.Tag = "iH1_Seban" Then Call Look_Hinbn(1): Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn(0): Exit Sub
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
        H1lb_Tornm = RV_TorRName
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
        H1lb_Tornm = RV_TorRName
        iH1_Ukeir = RV_Ukeir
    End If
End Sub
'+-----------------------------+
'+       ïiî‘åüçı              +
'+-----------------------------+
Private Sub Look_Hinbn(KB As Integer)
    RV_Left = 0
    RV_Top = 0
        
    RVI_Hinmoku = 1 ' î≠íç
    RVI_SK = 9
    RV_Call = ""
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        Select Case KB
            Case 1: iH1_Seban = RV_Seban
            Case Else
                iH1_Hinbn = RV_Hinbn
        End Select
    End If
End Sub
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
    ' îwî‘çÜÇÊÇËëºèÓïÒï\é¶
    If Trim(iH1_Seban) <> "" Then
        gSL_Select = "Select éÊà¯êÊ,éÛì¸,ïiî‘ From ïiî‘éÊà¯êÊÉ}ÉXÉ^"
        gSL_Select = gSL_Select & " Where îwî‘çÜ = '" & iH1_Seban & "'"
        gSL_Select = gSL_Select & "   And ïiñ⁄ = 1"
        If HTMRead(gSL_Select, 0, 1) Then
           If CisDB.RecordCount = 1 Then
              iH1_Torcd = HTM.éÊà¯êÊ
              iH1_Ukeir = HTM.éÛì¸
              iH1_Hinbn = HTM.ïiî‘
              Call iH1_Torcd_LostFocus
           End If
        End If
        Call HTMClose(1)
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
'+  iH1_Ymd
'+-----------------------------+
Private Sub iH1_Ymd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Ymd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'*******************************
'*      É^ÉCÉgÉãÉoÅ[ï“èW        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
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
'    PB_Chg.Visible = B1Mode
    
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

ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1¡™Ø∏èàóù
'****************************
Private Function Head1Chk() As Boolean
    Dim SetItem     As String
    Dim SetDayStr   As String
    Dim YSuryo      As Long
    Dim YYSuryo     As Long
    Dim wWhere      As String
    Dim wJEndDay    As Long
    
    Head1Chk = False
    
    If Trim(iH1_Ymd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    éwíËì˙ÅiîNåéñîÇÕîNåéì˙ÅjÇì¸óÕÇµÇƒâ∫Ç≥Ç¢         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_Ymd.SetFocus
        Exit Function
    End If
    If CisFun.Val2(iH1_Ymd.cYear) = 0 Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    éwíËì˙ÅiîNÅjÇì¸óÕÇµÇƒâ∫Ç≥Ç¢         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_Ymd.SetFocus
       Exit Function
    End If
    If CisFun.Val2(iH1_Ymd.cMonth) = 0 Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    éwíËì˙ÅiåéÅjÇì¸óÕÇµÇƒâ∫Ç≥Ç¢         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_Ymd.SetFocus
       Exit Function
    End If
    If CisFun.Val2(iH1_Ymd.cDay) <> 0 Then
       If Not iH1_Ymd.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    éwíËì˙Çê≥ÇµÇ¢ì˙ïtÇ…Çƒì¸óÕÇµÇƒâ∫Ç≥Ç¢         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_Ymd.SetFocus
        Exit Function
       End If
    End If
    
    Call GridInit
    
    Erase mSuryo
    Erase mMaisu
    Erase mNyuka
    
    PNL_Wait.Visible = True
    DoEvents
    
    With CisFun
         .Where_CreateKB = NewDocument
        ' î[ì¸ì˙
        .Where_Create AlphaMe, "Substring(î[ì¸ì˙,1,6)", Mid(iH1_Ymd, 1, 6)
        ' ïiî‘
        .Where_Create AlphaMe, "ïiî‘", iH1_Hinbn
        ' îwî‘çÜ
        .Where_Create AlphaMe, "îwî‘çÜ", iH1_Seban
        ' édì¸êÊ
        .Where_Create AlphaMe, "édì¸êÊ", iH1_Torcd
        ' éÛì¸
        .Where_Create AlphaMe, "éÛì¸", iH1_Ukeir
    
        wWhere = .Where_Phrase
    
    End With
    
    SetItem = ""
    If Trim(iH1_Seban) <> "" Then
       SetItem = "îwî‘çÜ,"
    End If
    If Trim(iH1_Hinbn) <> "" Then
       SetItem = SetItem & "ïiî‘,"
    End If
    If Trim(iH1_Torcd) <> "" Then
       SetItem = SetItem & "édì¸êÊ,"
    End If
    If Trim(iH1_Ukeir) <> "" Then
       SetItem = SetItem & "éÛì¸,"
    End If
       
    gSL_Select = "Select 2 åèêî," & SetItem & "î[ì¸ì˙,SUM(ñáêî) ñáêî,Sum(î≠íçêî) î≠íçêî"
    gSL_Select = gSL_Select & " From î≠íçÉeÅ[ÉuÉã "
    gSL_Select = gSL_Select & CisFun.Where_Phrase
    gSL_Select = gSL_Select & " And ï™î[ãÊï™ = 0"
    
    gSL_Select = gSL_Select & " Group By " & SetItem & "î[ì¸ì˙"
    
    gSL_Select = gSL_Select & " Union "
    
    gSL_Select = gSL_Select & "Select 3 åèêî," & SetItem & "ì¸â◊ì˙,SUM(ì¸â◊ñáêî) ñáêî,Sum(ì¸â◊êî) î≠íçêî"
    gSL_Select = gSL_Select & " From ì¸â◊ÉeÅ[ÉuÉã "
    gSL_Select = gSL_Select & Replace(wWhere, "î[ì¸ì˙", "ì¸â◊ì˙")
    gSL_Select = gSL_Select & " Group By " & SetItem & "ì¸â◊ì˙"
    
    gSL_Select = gSL_Select & " Union "
    
    wWhere = Replace(wWhere, "î[ì¸ì˙", "é¿ê—ì˙")
    wWhere = Replace(wWhere, "édì¸êÊ", "é¿ê—êÊ")
    SetItem = Replace(SetItem, "édì¸êÊ", "é¿ê—êÊ")
    
    gSL_Select = gSL_Select & "Select 6 åèêî," & SetItem & "é¿ê—ì˙,SUM(é¿ê—ñáêî) ñáêî,Sum(é¿ê—êî) î≠íçêî"
    gSL_Select = gSL_Select & " From é¿ê—ä«óùÉeÅ[ÉuÉã "
    gSL_Select = gSL_Select & wWhere
    gSL_Select = gSL_Select & " And é¿ê—ãÊï™ = 1"
    gSL_Select = gSL_Select & " Group By " & SetItem & "é¿ê—ì˙"
    
    gSL_Select = gSL_Select & " Union "
    
    gSL_Select = gSL_Select & "Select 4 åèêî," & SetItem & "é¿ê—ì˙,SUM(é¿ê—ñáêî) ñáêî,Sum(é¿ê—êî) î≠íçêî"
    gSL_Select = gSL_Select & " From é¿ê—ä«óùÉeÅ[ÉuÉã "
    gSL_Select = gSL_Select & wWhere
    gSL_Select = gSL_Select & " And é¿ê—ãÊï™ IN (2,3)"
    gSL_Select = gSL_Select & " Group By " & SetItem & "é¿ê—ì˙"
    
    If Not HCTRead(gSL_Select) Then
        Call HCTClose
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    äYìñÉfÅ[É^Ç™ë∂ç›ÇµÇ‹ÇπÇÒ°            "
            .MB_MSG(3) = "    çƒìxíäèoèåèÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢       "
            .MB_Title = "íäèoèàóù"
            .MB_Button = Error
            .MBOX
        End With
        iH1_Seban.SetFocus
        GoTo Head1Chk_Ed
    End If

    Do Until Not HCT_RDSTS
        
        gCnt2 = CisFun.Val2(Mid(HCT.î[ì¸ì˙, 7, 2))
        If gCnt2 <> 0 Then
           If HCT.åèêî = 6 Then
              mSuryo(3, gCnt2) = mSuryo(3, gCnt2) + HCT.î≠íçêî
              mMaisu(3, gCnt2) = mMaisu(3, gCnt2) + HCT.ñáêî
           Else
              mSuryo(HCT.åèêî, gCnt2) = mSuryo(HCT.åèêî, gCnt2) + HCT.î≠íçêî
              mMaisu(HCT.åèêî, gCnt2) = mMaisu(HCT.åèêî, gCnt2) + HCT.ñáêî
           End If
           If HCT.åèêî = 3 Then
              mNyuka(gCnt2) = mNyuka(gCnt2) + HCT.î≠íçêî
           End If
           ' é¿ê—ÇÃç≈èIì˙ÇãÅÇﬂÇÈ
           If wJEndDay < gCnt2 Then
              wJEndDay = gCnt2
           End If
        End If

        Call HCTReadNext
    Loop
    Call HCTClose
' ì‡é¶èÓïÒæØƒ
    gSL_Select = "Select HN.î[ä˙,SUM(HN.êîó ) êîó  From î≠íçì‡é¶ÉeÅ[ÉuÉã HN"
    gSL_Select = gSL_Select & " Left Outer Join ïiî‘éÊà¯êÊÉ}ÉXÉ^ HTM"
    gSL_Select = gSL_Select & "     ON  HN.ïiî‘ = HTM.ïiî‘"
    gSL_Select = gSL_Select & "     And HN.éËîzêÊ = HTM.éÊà¯êÊ"
    gSL_Select = gSL_Select & "     And HN.éÛì¸ = HTM.éÛì¸"
    With CisFun
         .Where_CreateKB = NewDocument
        ' ì‡é¶îNåé
        .Where_Create AlphaMe, "HN.ì‡é¶îNåé", Mid(iH1_Ymd, 1, 6)
        ' ïiî‘
        If Trim(iH1_Seban) = "" Then
           .Where_Create AlphaMe, "HN.ïiî‘", iH1_Hinbn, , Create
        Else
           .Where_Create AlphaMe, "HN.ïiî‘", iH1_Hinbn
        End If
        ' îwî‘çÜ
        .Where_Create AlphaMe, "HTM.îwî‘çÜ", iH1_Seban
        ' édì¸êÊ
        .Where_Create AlphaMe, "HN.éËîzêÊ", iH1_Torcd
        ' éÛì¸
        .Where_Create AlphaMe, "HN.éÛì¸", iH1_Ukeir
    
        ' ì‡é¶îNåé
        .Where_Create AlphaMe, "Substring(HN.î[ä˙,1,6)", Mid(iH1_Ymd, 1, 6)
        
        gSL_Select = gSL_Select & .Where_Phrase
    
    End With
    gSL_Select = gSL_Select & " Group By HN.î[ä˙"

    
    If HNJRead(gSL_Select) Then
        Do Until Not HNJ_RDSTS
            gCnt2 = CisFun.Val2(Mid(HNJ.î[ä˙, 7, 2))
            If gCnt2 <> 0 Then
               mSuryo(1, gCnt2) = mSuryo(1, gCnt2) + HNJ.êîó 
            End If
            
            Call HNJReadNext
        Loop
    End If
    Call HNJClose
    For gInt = 1 To 2
        gSL_Select = "Select SUM(HN.êîó ) êîó  From î≠íçì‡é¶ÉeÅ[ÉuÉã HN"
        gSL_Select = gSL_Select & " Left Outer Join ïiî‘éÊà¯êÊÉ}ÉXÉ^ HTM"
        gSL_Select = gSL_Select & "     ON  HN.ïiî‘ = HTM.ïiî‘"
        gSL_Select = gSL_Select & "     And HN.éËîzêÊ = HTM.éÊà¯êÊ"
        gSL_Select = gSL_Select & "     And HN.éÛì¸ = HTM.éÛì¸"
        With CisFun
             .Where_CreateKB = NewDocument
            ' ì‡é¶îNåé
            .Where_Create AlphaMe, "HN.ì‡é¶îNåé", Mid(iH1_Ymd, 1, 6)
            ' ïiî‘
            If Trim(iH1_Seban) = "" Then
               .Where_Create AlphaMe, "HN.ïiî‘", iH1_Hinbn, , Create
            Else
               .Where_Create AlphaMe, "HN.ïiî‘", iH1_Hinbn
            End If
            ' îwî‘çÜ
            .Where_Create AlphaMe, "HTM.îwî‘çÜ", iH1_Seban
            ' édì¸êÊ
            .Where_Create AlphaMe, "HN.éËîzêÊ", iH1_Torcd
            ' éÛì¸
            .Where_Create AlphaMe, "HN.éÛì¸", iH1_Ukeir
            
            .Date_RsFormat = fYM
            .Date_Add iH1_Ymd, M, gInt
            gStr = .Date_Result
            
            ' ì‡é¶îNåé
            .Where_Create AlphaMe, "Substring(î[ä˙,1,6)", Mid(gStr, 1, 6)
            
            gSL_Select = gSL_Select & .Where_Phrase
        
        End With
        If HNJRead(gSL_Select, 1) Then
           If gInt = 1 Then
              B1lb_YSyoyo = Format(HNJ.êîó , "#,##0")
           End If
           If gInt = 2 Then
              B1lb_YYSyoyo = Format(HNJ.êîó , "#,##0")
           End If
        End If
    Next gInt
    
    Call HNJClose
    ' ç›å…èÓïÒÉZÉbÉg
    Erase mZaiko
    Erase mSuryoK
    Erase mSuryoA
    
    If Trim(iH1_Hinbn) <> "" And InStr(1, iH1_Hinbn, "*") = 0 Then
               
       PB_Disp.Visible = True
       DispMei.Visible = True
       
       gSL_Select = "Select é¿íIì˙,ç›å…êî From ç›å…óöóÉ}ÉXÉ^ "
       gSL_Select = gSL_Select & " Where îNåéìx = '" & Mid(iH1_Ymd, 1, 6) & "'"
       gSL_Select = gSL_Select & "   And ïiî‘   = '" & iH1_Hinbn & "'"
       gSL_Select = gSL_Select & "   And óöóãÊï™ = 0"
       If ZKMRead(gSL_Select) Then
          Do Until Not ZKM_RDSTS
             gCnt2 = CisFun.Val2(Mid(ZKM.é¿íIì˙, 7, 2))
             If gCnt2 <> 0 Then
                mZaiko(gCnt2) = ZKM.ç›å…êî
             Else
                B1lb_TSyoyo(5) = Format(ZKM.ç›å…êî, "#,##0")
                mSuryo(5, gCnt2) = ZKM.ç›å…êî
                mSuryoK(gCnt2) = ZKM.ç›å…êî
                mSuryoA(gCnt2) = ZKM.ç›å…êî
             End If
             Call ZKMReadNext
          Loop
       End If
        
       Call ZKMClose
    
       gSL_Select = "Select é¿íIì˙,ç›å…êî From ç›å…É}ÉXÉ^ "
       gSL_Select = gSL_Select & " Where îNåéìx = '" & Mid(iH1_Ymd, 1, 6) & "'"
       gSL_Select = gSL_Select & "   And ïiî‘   = '" & iH1_Hinbn & "'"
       If ZKMRead(gSL_Select) Then
          mZaikoNew = CisFun.Val2(Mid(ZKM.é¿íIì˙, 7, 2))
          If mZaikoNew = 0 Then
             mSuryo(5, 1) = ZKM.ç›å…êî
             mZaiko(1) = ZKM.ç›å…êî
          Else
             mSuryo(5, mZaikoNew) = ZKM.ç›å…êî
             mZaiko(mZaikoNew) = ZKM.ç›å…êî
          End If
       End If
       
       For gCnt1 = 1 To 31
           If mZaiko(gCnt1) = 0 Or gCnt1 > mZaikoNew Then
              mSuryo(5, gCnt1) = mSuryo(5, gCnt1 - 1) + mSuryo(3, gCnt1) - mSuryo(4, gCnt1)
              mSuryoK(gCnt1) = mSuryoK(gCnt1 - 1) + mSuryo(2, gCnt1) - mSuryo(1, gCnt1)
              If gCnt1 > wJEndDay Then
                 mSuryoA(gCnt1) = mSuryoA(gCnt1 - 1) + mSuryo(2, gCnt1) - mSuryo(1, gCnt1)
              Else
                 mSuryoA(gCnt1) = mSuryoA(gCnt1 - 1) + mSuryo(3, gCnt1) - mSuryo(4, gCnt1)
              End If
           Else
              mSuryo(5, gCnt1) = mZaiko(gCnt1)
              mSuryoK(gCnt1) = mZaiko(gCnt1)
              mSuryoA(gCnt1) = mZaiko(gCnt1)
           End If
       Next gCnt1
        
       For gCnt1 = 1 To 31
           If mZaiko(gCnt1) <> 0 Then
              If gCnt1 <= mZaikoNew And mZaikoNew > 1 Then
                 For gCnt2 = gCnt1 - 1 To 1 Step -1
                     mSuryo(5, gCnt2) = mSuryo(5, gCnt2 + 1) + mSuryo(4, gCnt2 + 1) - mSuryo(3, gCnt2 + 1)
                     mSuryoK(gCnt2) = mSuryoK(gCnt2 + 1) + mSuryo(2, gCnt2 + 1) - mSuryo(1, gCnt2 + 1)
                     If gCnt1 > wJEndDay Then
                        mSuryoA(gCnt2) = mSuryoA(gCnt2 + 1) + mSuryo(2, gCnt2 + 1) - mSuryo(1, gCnt2 + 1)
                     Else
                        mSuryoA(gCnt2) = mSuryoA(gCnt2 + 1) + mSuryo(3, gCnt2 + 1) - mSuryo(4, gCnt2 + 1)
                     End If
                 Next gCnt2
                 Exit For
              End If
           End If
       Next gCnt1
    Else
       PB_Disp.Visible = False
       DispMei.Visible = False
    End If
    
    Call GridValueSet
    
    With vsGrid
        If CisFun.Val2(iH1_Ymd.cDay) = 0 Then
           If Mid(iH1_Ymd, 1, 6) <> Format(Date, "YYYYMM") Then
              gCnt1 = 1
           Else
              gCnt1 = CisFun.Val2(Format(Date, "dd"))
           End If
        Else
           gCnt1 = CisFun.Val2(iH1_Ymd.cDay)
        End If
        mCellSet = True
        .Row = mDayRow(gCnt1)
        .Col = mDayCol(gCnt1)
        Call MeisaiDisp
        mCellSet = False
    End With
    
' * ïiî‘éÊà¯êÊ
    gSL_Select = "SELECT * "
    gSL_Select = gSL_Select & " FROM ïiî‘éÊà¯êÊÉ}ÉXÉ^  "
    gSL_Select = gSL_Select & " WHERE ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    If Trim(iH1_Torcd) <> "" Then
       gSL_Select = gSL_Select & "   AND éÊà¯êÊ = '" & RTrim(iH1_Torcd) & "'"
    End If
    If Trim(iH1_Ukeir) <> "" Then
       gSL_Select = gSL_Select & "   AND éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    End If
    gSL_Select = gSL_Select & "   AND ïiñ⁄ = 1 "
    If HTMRead(gSL_Select) Then
        With HTM
            If CisDB.RecordCount = 1 Then
                ' * SKãÊï™
                If .SKãÊï™ = 0 Then
                    B3lb_SKKbn = "Ç©ÇÒÇŒÇÒ"
                Else
                    B3lb_SKKbn = "éw é¶"
                End If
                ' * ÉTÉCÉNÉã
                If RTrim(.ÉTÉCÉNÉã) <> "" Then
                    B3lb_Cycle = Mid(.ÉTÉCÉNÉã, 1, 1) & "-" & Mid(.ÉTÉCÉNÉã, 2, 2) & "-" & Mid(.ÉTÉCÉNÉã, 4, 2)
                End If
                ' ÉäÅ[ÉhÉ^ÉCÉÄ
                B3lb_LTime = CisFun.RSetFld(HTM.LT, 2, "##")
                ' * é˚óeêî1
                If .é˚óeêî1 <> 0 Then: B3lb_Syuyo1 = CisFun.RSetFld(.é˚óeêî1, 7, "#,###")
                ' * é˚óeêî2
                If .é˚óeêî2 <> 0 Then: B3lb_Syuyo2 = CisFun.RSetFld(.é˚óeêî2, 7, "#,###")
                ' * êÿë÷ì˙ÅEï÷
                If RTrim(.ïœçXì˙) <> "" Then
                    B3lb_CDate = Mid(.ïœçXì˙, 1, 4) & "/" & Mid(.ïœçXì˙, 5, 2) & "/" & Mid(.ïœçXì˙, 7, 2)
                End If
                If .ïœçXï÷ <> "" Then: B3lb_CBin = CisFun.RSetFld(.ïœçXï÷, 2, "00")
                ' î≠íçÉçÉbÉg1
                B3lb_Lot1 = CisFun.RSetFld(.ÉçÉbÉg1, 6, "#,###")
                ' î≠íçÉçÉbÉg2
                B3lb_Lot2 = CisFun.RSetFld(.ÉçÉbÉg2, 6, "#,###")
                ' ójì˙
                If .Week1 <> 0 Or .Week2 <> 0 Or _
                   .Week3 <> 0 Or .Week4 <> 0 Or _
                   .Week5 <> 0 Then
                   Back_Week = "ójì˙éwíË"
                    If .Week1 <> 0 Then: B3lb_Week(1) = "åé"
                    If .Week2 <> 0 Then: B3lb_Week(2) = "âŒ"
                    If .Week3 <> 0 Then: B3lb_Week(3) = "êÖ"
                    If .Week4 <> 0 Then: B3lb_Week(4) = "ñÿ"
                    If .Week5 <> 0 Then: B3lb_Week(5) = "ã‡"
                End If
                ' * ì˙ït
                If .Day1 <> 0 Or .Day2 <> 0 Or _
                   .Day3 <> 0 Or .Day4 <> 0 Or _
                   .Day5 <> 0 Then
                   Back_Week = "ì˙ïtéwíË"
                   If .Day1 <> 0 Then: B3lb_Week(1) = CisFun.RSetFld(.Day1, 2, "##")
                   If .Day2 <> 0 Then: B3lb_Week(2) = CisFun.RSetFld(.Day2, 2, "##")
                   If .Day3 <> 0 Then: B3lb_Week(3) = CisFun.RSetFld(.Day3, 2, "##")
                   If .Day4 <> 0 Then: B3lb_Week(4) = CisFun.RSetFld(.Day4, 2, "##")
                   If .Day5 <> 0 Then: B3lb_Week(5) = CisFun.RSetFld(.Day5, 2, "##")
                End If
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
                    B3lb_ZKKbn = HNM.ç›å…ä«óùãÊï™ñº
                    Select Case HNM.ç›å…ä«óùãÊï™
                        Case 2: B3lb_ZKSu = Format(HNM.ç›å…ä«óùêîíl, "###") & " Åì"
                        Case 3: B3lb_ZKSu = Format(HNM.ç›å…ä«óùêîíl, "###,###")
                        Case 4: B3lb_ZKSu = Format(HNM.ç›å…ä«óùêîíl, "#0.0") & " ì˙"
                        Case Else: B3lb_ZKSu = ""
                    End Select
                End If
            End If
        End With
    End If
    Call HTMClose
    
    Head1Chk = True

Head1Chk_Ed:
    PNL_Wait.Visible = False
End Function
'**************************************
'*                                    *
'*          ÉOÉäÉbÉhèàóù              *
'*                                    *
'**************************************
Private Sub GridInit()
    Dim SetColor        As Boolean
    
    mCellSet = True
    
    With vsGrid
        .Redraw = False
        .Rows = 1: .Cols = 1
        .Rows = 36: .Cols = 9
        
        .Row = 0
        For gCnt2 = 0 To .Cols - 1
            .Col = gCnt2
            .ColAlignment(gCnt2) = flexAlignCenterCenter
            If gCnt2 > 0 Then
                If gCnt2 = 1 Then
                   .ColWidth(gCnt2) = 30
                Else
                   .ColWidth(gCnt2) = 1300
                End If
            End If
        Next gCnt2
        
        Call DaySet
        
        For gCnt1 = 0 To .Rows - 1
            .Row = gCnt1: .Col = 0
            If gCnt1 Mod 6 = 0 Then
                .Text = "ì˙ ït"
                For gCnt2 = 0 To .Cols - 1
                    .Col = gCnt2
                Next gCnt2
            Else
                .Text = Grid_RowHead(gCnt1 Mod 6).Tag
                .CellBackColor = B1lb_TSyoyo(gCnt1 Mod 6).BackColor
            End If

            For gCnt2 = 2 To .Cols - 1
                .Col = gCnt2
                SetColor = True
                If gCnt1 Mod 6 = 0 Then
                   SetColor = False
                End If
                If gCnt1 < 6 Then
                   If Trim(.TextMatrix(0, gCnt2)) = "" Then
                      .CellBackColor = DayColor.BackColor
                      SetColor = False
                   End If
                End If
                If gCnt1 > 24 Then
                   If Trim(.TextMatrix(24, gCnt2)) = "" Then
                      .CellBackColor = DayColor.BackColor
                      SetColor = False
                   End If
                End If
                If gCnt1 > 30 Then
                   If Trim(.TextMatrix(30, gCnt2)) = "" Then
                      .CellBackColor = DayColor.BackColor
                      SetColor = False
                   End If
                End If
                If SetColor Then
                   .CellBackColor = B1lb_TSyoyo(gCnt1 Mod 6).BackColor
                   .CellAlignment = flexAlignRightCenter
                End If
                
            Next gCnt2
        Next gCnt1
        
        .Row = 1: .Col = 2
        .Redraw = True
    End With
    mCellSet = False

End Sub

Private Sub DaySet()
    Dim wDayCol         As Integer
    Dim SetDay          As Integer
    Dim EndDay          As Integer
    Dim wKado           As Boolean
    
    wDayCol = 1
    SetDay = 1
    
'       *  ÉJÉåÉìÉ_ïîì˙ït ÉZÉbÉg  *
    gSL_Select = "Select * From ÉJÉåÉìÉ_É}ÉXÉ^"
    gSL_Select = gSL_Select & " Where éÊà¯êÊãÊï™ = 1"
    gSL_Select = gSL_Select & " And   éÊà¯êÊ = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   îN = '" & Format(iH1_Ymd.cYear, "0000") & "'"
    gSL_Select = gSL_Select & " And   åé = '" & Format(iH1_Ymd.cMonth, "00") & "'"
    gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    If Not CLMRead(gSL_Select, 1) Then
'       gSL_Select = "Select * From ÉJÉåÉìÉ_É}ÉXÉ^"
'       gSL_Select = gSL_Select & " Where éÊà¯êÊãÊï™ = 1"
'       gSL_Select = gSL_Select & " And   éÊà¯êÊ = '" & RTrim(iH1_Torcd) & "'"
'       gSL_Select = gSL_Select & " And   îN = '" & Format(iH1_YM.cYear, "0000") & "'"
'       gSL_Select = gSL_Select & " And   åé = '" & Format(iH1_YM.cMonth, "00") & "'"
'       gSL_Select = gSL_Select & " And   éÛì¸ = ''"
'       If Not CLMRead(gSL_Select, 1) Then
'          gSL_Select = "Select * From ÉJÉåÉìÉ_É}ÉXÉ^"
'          gSL_Select = gSL_Select & " Where éÊà¯êÊãÊï™ = 1"
'          gSL_Select = gSL_Select & " And   éÊà¯êÊ = ''"
'          gSL_Select = gSL_Select & " And   îN = '" & Format(iH1_YM.cYear, "0000") & "'"
'          gSL_Select = gSL_Select & " And   åé = '" & Format(iH1_YM.cMonth, "00") & "'"
'          gSL_Select = gSL_Select & " And   éÛì¸ = ''"
'          If Not CLMRead(gSL_Select, 1) Then
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
                 .Date_End iH1_Ymd
                 CLM.ññì˙ = .Val2(.Date_Result)
                 CLM.â“ìÆì˙êî = CLM.ññì˙
             End With
'          End If
'        End If
    End If
    
    EndDay = CLM.ññì˙
    gDate = DateValue(iH1_Ymd.cYear & "/" & iH1_Ymd.cMonth & "/01")
    gInt = Weekday(gDate, mWeekStart)
    Erase mDayRow
    Erase mDayCol
    
    With vsGrid
        For gCnt1 = 0 To .Rows - 1 Step 6
            .Row = gCnt1
            For gCnt2 = 1 To .Cols - 2
                If gCnt2 >= gInt Then
                   .Col = gCnt2 + 1
                   If SetDay <= EndDay Then
                      gDate = DateValue(iH1_Ymd.cYear & "/" & iH1_Ymd.cMonth & "/" & SetDay)
                      gCnt3 = Weekday(gDate, mWeekStart)
                      .CellBackColor = H1lb_Tornm.BackColor
                      
                      wKado = False
                      If Mid(CLM.â“ìÆãÊï™, SetDay, 1) = "0" Or _
                         Mid(CLM.ñÈãŒãÊï™, SetDay, 1) = "0" Then
                         wKado = True
                      End If
                      If Not wKado Then
                         .CellBackColor = ColorSunday.BackColor
                      End If
                      
                      If mWeekStart = 2 And gCnt3 = 7 And Not wKado Then
                         .CellBackColor = ColorSunday.BackColor
                      End If
                      If mWeekStart = 2 And gCnt3 = 6 And Not wKado Then
                         .CellBackColor = ColorSata.BackColor
                      End If
                      If mWeekStart = 1 And gCnt3 = 1 And Not wKado Then
                         .CellBackColor = ColorSunday.BackColor
                      End If
                      If mWeekStart = 1 And gCnt3 = 7 And Not wKado Then
                         .CellBackColor = ColorSata.BackColor
                      End If
                      
                      If Date = gDate And CisFun.Val2(iH1_Ymd.cDay) = 0 Then
                         .CellBackColor = ColorToday.BackColor
                      End If
                      If CisFun.Val2(iH1_Ymd.cDay) <> 0 And CisFun.Val2(iH1_Ymd.cDay) = SetDay Then
                         .CellBackColor = ColorToday.BackColor
                      End If
                      
                      .TextMatrix(gCnt1, gCnt2 + 1) = SetDay & " (" & WeekdayName(gCnt2, True, mWeekStart) & ")"
                      mDayRow(SetDay) = gCnt1
                      mDayCol(SetDay) = gCnt2 + 1

                      SetDay = SetDay + 1
                      gInt = 0
                   Else
                      .CellBackColor = DayColor.BackColor
                   End If
                Else
                    .CellBackColor = DayColor.BackColor
                End If
            Next gCnt2
        Next gCnt1
    End With
End Sub
Private Sub GridValueSet()
    Dim TSuryo(5)    As Long
    Dim SSuryo(5)    As Long
    Dim SetSu        As Long
    
    mCellSet = True
    
    With vsGrid
        .Redraw = False
        For gCnt1 = 1 To 31
            If mDayRow(gCnt1) = 0 And mDayCol(gCnt1) = 0 Then Exit For
            gDate = DateValue(iH1_Ymd.cYear & "/" & iH1_Ymd.cMonth & "/" & gCnt1)
            .Col = mDayCol(gCnt1)
            For gCnt2 = 1 To 5
                If mSetType = 0 Then
                    SetSu = mSuryo(gCnt2, gCnt1)
                    If gCnt2 = 5 Then
                       If mSetDisp = 1 Then
                          SetSu = mSuryoK(gCnt1)
                       End If
                       If mSetDisp = 2 Then
                          SetSu = mSuryoA(gCnt1)
                       End If
                    End If
                    .TextMatrix(mDayRow(gCnt1) + gCnt2, mDayCol(gCnt1)) = _
                                            Format(SetSu, "#,###")
                    If gCnt2 = 5 Then
                       If mSuryo(1, gCnt1) = 0 And _
                          mSuryo(2, gCnt1) = 0 And _
                          mSuryo(3, gCnt1) = 0 And _
                          mSuryo(4, gCnt1) = 0 Then
                          .TextMatrix(mDayRow(gCnt1) + gCnt2, mDayCol(gCnt1)) = ""
                       End If
                       If mZaiko(gCnt1) <> 0 Then
                          If gCnt1 = mZaikoNew Then
                             .TextMatrix(mDayRow(gCnt1) + gCnt2, mDayCol(gCnt1)) = _
                                       "**" & CisFun.RSetFld(SetSu, 8, "#,###")
                          Else
                             .TextMatrix(mDayRow(gCnt1) + gCnt2, mDayCol(gCnt1)) = _
                                       "*" & CisFun.RSetFld(SetSu, 9, "#,###")
                          End If
                       End If
                    End If
                Else
                    SetSu = mMaisu(gCnt2, gCnt1)
                    .TextMatrix(mDayRow(gCnt1) + gCnt2, mDayCol(gCnt1)) = _
                                            Format(SetSu, "#,###")
                End If
                .Row = mDayRow(gCnt1) + gCnt2
                If Date = gDate And CisFun.Val2(iH1_Ymd.cDay) = 0 Then
                   .CellBackColor = ColorToday.BackColor
                End If
                If CisFun.Val2(iH1_Ymd.cDay) <> 0 And CisFun.Val2(iH1_Ymd.cDay) = gCnt1 Then
                   .CellBackColor = ColorToday.BackColor
                End If
                
                TSuryo(gCnt2) = TSuryo(gCnt2) + SetSu
                
                If Date >= gDate And CisFun.Val2(iH1_Ymd.cDay) = 0 Then
                   SSuryo(gCnt2) = SSuryo(gCnt2) + SetSu
                End If
                If CisFun.Val2(iH1_Ymd.cDay) <> 0 And CisFun.Val2(iH1_Ymd.cDay) >= gCnt1 Then
                   SSuryo(gCnt2) = SSuryo(gCnt2) + SetSu
                End If
                
            Next gCnt2
        Next gCnt1
        
        For gCnt1 = 1 To 4
            B1lb_TSyoyo(gCnt1) = Format(TSuryo(gCnt1), "#,###")
            B1lb_SSuryo(gCnt1) = Format(SSuryo(gCnt1), "#,###")
        Next gCnt1
        .Redraw = True
    End With
    
    mCellSet = False
    
End Sub
Private Function GetDay() As Integer
    Dim wCnt        As Long
    GetDay = 0
    With vsGrid
        For wCnt = 1 To 31
            If mDayRow(wCnt) = (.Row \ 6) * 6 And _
               mDayCol(wCnt) = .Col Then
               GetDay = wCnt
               Exit Function
            End If
        Next wCnt
    End With
End Function

Private Sub MeisaiDisp()
    Dim SetItem     As String
    Dim wWhere      As String
    Dim wOrder      As String
    Dim SetDay      As Integer
    Dim SetType     As Integer
    
    SetDay = GetDay
    SetType = vsGrid.Row Mod 6
    
    If SetDay = 0 Or InStr(1, vsGrid.TextMatrix(mDayRow(SetDay) + 5, vsGrid.Col), "*") = 0 Then
       Call DBGridClear
       H3_Area1.Visible = False
       B3_Area2.Visible = True
       Exit Sub
    End If
    
    With CisFun
         .Where_CreateKB = NewDocument
        ' îNåéìx
        .Where_Create AlphaMe, "îNåéìx", Mid(iH1_Ymd, 1, 6)
        ' é¿íIì˙
        .Where_Create AlphaMe, "é¿íIì˙", Mid(iH1_Ymd, 1, 6) & Format(SetDay, "00")
        ' ïiî‘
        .Where_Create AlphaMe, "ïiî‘", iH1_Hinbn
    
        wWhere = .Where_Phrase
    
    End With
    
    gSL_Select = "Select ZM.ì¸óÕì˙ït,ZM.ç›å…êî,ZM.ïœçXéñóR,RTrim(SY.ílñºèÃ) íSìñé“ñº"
    gSL_Select = gSL_Select & " From ç›å…óöóÉ}ÉXÉ^ ZM"
    gSL_Select = gSL_Select & " Left Outer Join ñºèÃÉ}ÉXÉ^ SY"
    gSL_Select = gSL_Select & "     On SY.ãÊï™ñºèÃ = 'íSìñä«óù'"
    gSL_Select = gSL_Select & "     And SY.íl = ZM.ì¸óÕíSìñé“"
    gSL_Select = gSL_Select & CisFun.Where_Phrase
    
    gSL_Select = gSL_Select & " Order By ZM.ì¸óÕì˙ït Desc"
    Adodc1.RecordSource = gSL_Select
    
    Adodc1.Refresh
    
    If Adodc1.Recordset.RecordCount > 0 Then
        H3_Area1.Visible = True
        B3_Area2.Visible = False
    End If

End Sub
Private Sub DBGridClear()
    DBGrid1.Close True
    On Error Resume Next
    Adodc1.Recordset.Close
    On Error GoTo 0
End Sub
Private Sub vsGrid_RowColChange()
    If mCellSet Then Exit Sub
    Call MeisaiDisp
End Sub
Private Sub vsGrid_DblClick()
    Dim SetDay      As Integer
    Dim SetType     As Integer
    Dim CallFlg     As Boolean
    
    
    SetDay = GetDay
    SetType = vsGrid.Row Mod 6
    
    CallFlg = False
    With vsGrid
        If mSuryo(2, SetDay) <> 0 Or _
           mNyuka(SetDay) <> 0 Then
           CallFlg = True
        End If
    End With
    
    If SetDay = 0 Or Not CallFlg Then
       Exit Sub
    End If
        
    Call CisFrm.MousePT(1)
    
    RV_SYmd = Mid(iH1_Ymd, 1, 6) & Format(SetDay, "00")
    RV_Hinbn = RTrim(iH1_Hinbn)
    RV_Seban = RTrim(iH1_Seban)
    RV_Torcd = RTrim(iH1_Torcd)
    RV_Ukeir = RTrim(iH1_Ukeir)
    Rv_Call3 = "K"
    Rv_CallRtn = 0
    If SetType = 3 Then
       Rv_CallRtn = 1
    End If
    
    CXH0420.Hide
    
    Set Rv_CallObj = CXH0420
    Set Rv_CallObj2 = CXH0425
    
    CXH0425.Show
    
    
    CXH0410.Show
'    Unload CXH0410
'    Set CXH0410 = Nothing
'
'    Call CisFrm.MousePT(11)
'
'    CXH0420.Show
End Sub

