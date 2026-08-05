VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXM0030 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "ïiî‘É}ÉXÉ^ÉÅÉìÉeÉiÉìÉX"
   ClientHeight    =   11160
   ClientLeft      =   2550
   ClientTop       =   1245
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z µ∞¿ﬁ∞
   ScaleHeight     =   11160
   ScaleWidth      =   15360
   WindowState     =   2  'ç≈ëÂâª
   Begin VB.PictureBox B3_Area3 
      Height          =   735
      Left            =   7380
      ScaleHeight     =   675
      ScaleWidth      =   3315
      TabIndex        =   111
      Top             =   5400
      Width           =   3375
      Begin Cis3D_v60.CIS3D PB_UTN 
         Height          =   675
         Left            =   0
         Top             =   15
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1191
         BackColor       =   16761087
         Caption         =   "îÑè„íPâø(F4)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   6
         cButton         =   -1  'True
         cPositionY      =   65
      End
      Begin Cis3D_v60.CIS3D PB_KTN 
         Height          =   675
         Left            =   1110
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1191
         BackColor       =   16761024
         Caption         =   "çwì¸íPâø(F5)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   6
         cButton         =   -1  'True
         cPositionY      =   65
      End
      Begin Cis3D_v60.CIS3D PB_STN 
         Height          =   675
         Left            =   2220
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1191
         BackColor       =   12632319
         Caption         =   "éxããíPâø(F6)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   6
         cButton         =   -1  'True
         cPositionY      =   65
      End
   End
   Begin VB.PictureBox B3_Area1 
      Height          =   4920
      Left            =   120
      ScaleHeight     =   4860
      ScaleWidth      =   10620
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   6180
      Width           =   10680
      Begin Cis3D_v60.CIS3D Back_HachuPHA 
         Height          =   510
         Left            =   75
         Top             =   4305
         Width           =   10440
         _ExtentX        =   18415
         _ExtentY        =   900
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
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisText_V60.CisText iB3_AtoCD 
            Height          =   330
            Left            =   7440
            TabIndex        =   118
            Top             =   120
            Width           =   480
            _ExtentX        =   847
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
            Text            =   "XXX"
            MaxLength       =   3
         End
         Begin Cis3D_v60.CIS3D B3lb_AtoNM 
            Height          =   330
            Left            =   7905
            Top             =   120
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXX"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
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
         Begin Cis3D_v60.CIS3D CIS3D18 
            Height          =   300
            Index           =   0
            Left            =   6600
            Top             =   150
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "å„çHíˆ"
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
            cBoderStyle     =   2
         End
         Begin CisText_V60.CisText iB3_Point 
            Height          =   360
            Left            =   1200
            TabIndex        =   116
            Top             =   120
            Width           =   330
            _ExtentX        =   582
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
         Begin Cis3D_v60.CIS3D B3lb_HenNM 
            Height          =   330
            Left            =   3945
            Top             =   120
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXX"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
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
         Begin CisText_V60.CisText iB3_HenCD 
            Height          =   330
            Left            =   3480
            TabIndex        =   117
            Top             =   120
            Width           =   480
            _ExtentX        =   847
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
            Text            =   "XXX"
            MaxLength       =   3
         End
         Begin Cis3D_v60.CIS3D CIS3D18 
            Height          =   300
            Index           =   1
            Left            =   2400
            Top             =   150
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "ï‘ãpèÍèä"
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
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D CIS3D18 
            Height          =   300
            Index           =   2
            Left            =   120
            Top             =   150
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "É|ÉCÉìÉg"
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
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D CIS3D13 
            Height          =   435
            Left            =   1560
            Top             =   30
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   767
            BackColor       =   16777152
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
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
            Begin VB.Image Image1 
               Height          =   435
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   435
            End
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   4
         Left            =   1515
         Top             =   1605
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "î[ì¸êÊ(íºëó)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB3_NTorcd 
            Height          =   360
            Left            =   60
            TabIndex        =   45
            Top             =   360
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
         Begin Cis3D_v60.CIS3D B3lb_NTorcd 
            Height          =   360
            Left            =   1050
            Top             =   360
            Width           =   1890
            _ExtentX        =   3334
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
         Begin CisText_V60.CisText iB3_Ukeir 
            Height          =   390
            Left            =   2970
            TabIndex        =   46
            Top             =   345
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   688
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
         Begin Cis3D_v60.CIS3D Back_Ukeir 
            Height          =   465
            Left            =   2910
            Top             =   30
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   820
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
            cAlingnment     =   6
            cBoderStyle     =   2
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   420
         Index           =   18
         Left            =   75
         Top             =   3885
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   741
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   420
            Left            =   5550
            Top             =   0
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   741
            ForeColor       =   16711680
            Caption         =   "çÜï‚"
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
            Begin VB.CheckBox B3Chk_Gouho 
               BackColor       =   &H00C0C0C0&
               Caption         =   "çÜå˚"
               Height          =   270
               Left            =   660
               TabIndex        =   115
               Top             =   90
               Width           =   795
            End
         End
         Begin VB.OptionButton B3Op_HinKu 
            BackColor       =   &H00C0C0C0&
            Caption         =   "äÆê¨ïi"
            Height          =   240
            Index           =   0
            Left            =   60
            TabIndex        =   81
            Top             =   90
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton B3Op_HinKu 
            BackColor       =   &H00C0C0C0&
            Caption         =   "çwì¸ïi"
            Height          =   240
            Index           =   1
            Left            =   1080
            TabIndex        =   82
            Top             =   90
            Width           =   1035
         End
         Begin VB.OptionButton B3Op_HinKu 
            BackColor       =   &H00C0C0C0&
            Caption         =   "édä|ïi"
            Height          =   240
            Index           =   2
            Left            =   2115
            TabIndex        =   83
            Top             =   90
            Width           =   1035
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   420
            Index           =   17
            Left            =   7020
            Top             =   0
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   741
            ForeColor       =   16711680
            Caption         =   "ëOçHíˆ"
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
            cPositionX      =   30
            Begin VB.CheckBox B3Chk_MaeHak 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ï\é¶ñ≥"
               Height          =   240
               Left            =   780
               TabIndex        =   114
               Top             =   105
               Width           =   1035
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   420
            Index           =   38
            Left            =   3180
            Top             =   0
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   741
            ForeColor       =   16711680
            Caption         =   "äÓèÄíPâø"
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
            Begin CisText_V60.CisText iB3_Kijyun 
               Height          =   360
               Left            =   1140
               TabIndex        =   84
               Top             =   30
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   635
               cFormat         =   "##,##0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "##,##0.00"
               MaxLength       =   9
               cDataType       =   2
               cDataReplace    =   1
               cFcDecimal      =   2
               cFdAutoFormat   =   1
               cGFormat        =   "####0.00"
               cILength        =   5
            End
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   420
         Index           =   21
         Left            =   8955
         Top             =   3885
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   741
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
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   420
            Index           =   32
            Left            =   0
            Top             =   0
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   741
            ForeColor       =   16711680
            Caption         =   "íºì`"
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
            Begin VB.CheckBox B3Chk_PrtkbSN1 
               BackColor       =   &H00C0C0C0&
               Caption         =   "î≠çsóL"
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   600
               TabIndex        =   113
               Top             =   75
               Width           =   930
            End
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   750
         Index           =   16
         Left            =   3645
         Top             =   3135
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   1323
         ForeColor       =   16711680
         Caption         =   "éÛóÃèë"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.CheckBox B3Chk_PrtkbJ 
            BackColor       =   &H00C0C0C0&
            Caption         =   "î≠çsóL"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   90
            TabIndex        =   75
            Top             =   405
            Width           =   960
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   750
         Index           =   14
         Left            =   1515
         Top             =   3135
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   1323
         ForeColor       =   16711680
         Caption         =   "î[ïièë"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.CheckBox B3Chk_PrtkbN 
            BackColor       =   &H00C0C0C0&
            Caption         =   "î≠çsóL"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   60
            TabIndex        =   73
            Top             =   390
            Width           =   945
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   750
         Index           =   15
         Left            =   2580
         Top             =   3135
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   1323
         ForeColor       =   16711680
         Caption         =   "éxããèë"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.CheckBox B3Chk_PrtkbS 
            BackColor       =   &H00C0C0C0&
            Caption         =   "î≠çsóL"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   75
            TabIndex        =   74
            Top             =   405
            Width           =   960
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   3
         Left            =   75
         Top             =   1605
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   1349
         ForeColor       =   16711680
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
         cPositionY      =   30
         Begin VB.ComboBox B3Comb_Tehai 
            BackColor       =   &H00C0E0FF&
            Height          =   360
            Left            =   120
            Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
            TabIndex        =   44
            Top             =   360
            Width           =   1170
         End
         Begin Cis3D_v60.CIS3D B3lb_Tehai 
            Height          =   330
            Left            =   120
            Top             =   360
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "é–ì‡"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   750
         Index           =   13
         Left            =   8310
         Top             =   3135
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   1323
         ForeColor       =   16711680
         Caption         =   "PNO"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB3_PNo 
            Height          =   360
            Left            =   90
            TabIndex        =   79
            Top             =   360
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
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
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   1
         Left            =   5775
         Top             =   1605
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1349
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisText_V60.CisText iB3_Hiri 
            Height          =   360
            Left            =   30
            TabIndex        =   49
            Top             =   360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   635
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
            cFbComma        =   0
            cFdAutoFormat   =   1
            cGFormat        =   "###"
            cILength        =   3
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "%"
            Height          =   315
            Left            =   540
            TabIndex        =   96
            Top             =   420
            Width           =   165
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   0
         Left            =   4965
         Top             =   1605
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "Ç©åWêî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB3_KKeisu 
            Height          =   360
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   635
            cFormat         =   "0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0.00"
            MaxLength       =   4
            cDataType       =   2
            cDataReplace    =   1
            cFbComma        =   0
            cFcDecimal      =   2
            cFdAutoFormat   =   1
            cGFormat        =   "0.00"
            cILength        =   1
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   750
         Index           =   11
         Left            =   4710
         Top             =   3135
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1323
         ForeColor       =   16711680
         Caption         =   "éÛóÃèëî≠çs"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.CheckBox B3Chk_Juryo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "î≠íçéû"
            Height          =   240
            Left            =   105
            TabIndex        =   76
            Top             =   420
            Width           =   1080
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   750
         Index           =   5
         Left            =   75
         Top             =   3135
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   1323
         ForeColor       =   16711680
         Caption         =   "ì`ï[éÌóﬁ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB3_DenSyu 
            Height          =   345
            Left            =   75
            TabIndex        =   72
            Top             =   360
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
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
         Begin Cis3D_v60.CIS3D B3lb_DenSyu 
            Height          =   330
            Left            =   315
            Top             =   360
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   582
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   750
         Index           =   12
         Left            =   5955
         Top             =   3135
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   1323
         ForeColor       =   16711680
         Caption         =   "ñæç◊èëî≠çs"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.ComboBox B3Comb_Meaisai 
            BackColor       =   &H00C0E0FF&
            Height          =   360
            Left            =   120
            Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
            TabIndex        =   77
            Top             =   360
            Width           =   1335
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   8
         Left            =   6495
         Top             =   1605
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "ì‡é¶î≠íç"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.CheckBox B3Chk_Naiji 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ÇµÇ»Ç¢"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   105
            TabIndex        =   50
            Top             =   420
            Width           =   930
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   7
         Left            =   3855
         Top             =   2370
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "ÉçÉbÉgáA"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB3_LSu2 
            Height          =   360
            Left            =   90
            TabIndex        =   61
            Top             =   360
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   635
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
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   750
         Index           =   6
         Left            =   8835
         Top             =   3135
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1323
         ForeColor       =   16711680
         Caption         =   "îı  çl"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB3_Biko 
            Height          =   360
            Left            =   60
            TabIndex        =   80
            Top             =   360
            Width           =   1560
            _ExtentX        =   2752
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   30
            IMEMode         =   4
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   750
         Index           =   31
         Left            =   7515
         Top             =   3135
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1323
         ForeColor       =   16711680
         Caption         =   "çHêî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB3_Kosu 
            Height          =   360
            Left            =   30
            TabIndex        =   78
            Top             =   360
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   635
            cFormat         =   "##0.0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "##0.0"
            MaxLength       =   5
            cDataType       =   2
            cDataReplace    =   1
            cFbComma        =   0
            cFcDecimal      =   1
            cFdAutoFormat   =   1
            cGFormat        =   "##0.0"
            cILength        =   3
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   20
         Left            =   9795
         Top             =   1605
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "åWêî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB3_HKeisu 
            Height          =   360
            Left            =   90
            TabIndex        =   55
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   635
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
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "%"
            Height          =   315
            Left            =   510
            TabIndex        =   97
            Top             =   450
            Width           =   165
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   19
         Left            =   7650
         Top             =   1605
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "ïΩ  èÄ  âª"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.OptionButton B3Op_Heijun 
            BackColor       =   &H00C0C0C0&
            Caption         =   "åé"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   3
            Left            =   1605
            TabIndex        =   54
            Top             =   360
            Width           =   450
         End
         Begin VB.OptionButton B3Op_Heijun 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ï÷"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   1095
            TabIndex        =   53
            Top             =   360
            Width           =   450
         End
         Begin VB.OptionButton B3Op_Heijun 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ì˙"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   585
            TabIndex        =   52
            Top             =   360
            Width           =   450
         End
         Begin VB.OptionButton B3Op_Heijun 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ñ≥"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   75
            TabIndex        =   51
            Top             =   360
            Value           =   -1  'True
            Width           =   450
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   29
         Left            =   2745
         Top             =   2370
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "ÉçÉbÉgá@"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB3_LSu 
            Height          =   360
            Left            =   90
            TabIndex        =   60
            Top             =   360
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   635
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
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   24
         Left            =   8310
         Top             =   2370
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "ì˙ ït éw íË"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB3_Day 
            Height          =   360
            Index           =   0
            Left            =   120
            TabIndex        =   67
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   635
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
         Begin CisText_V60.CisText iB3_Day 
            Height          =   360
            Index           =   1
            Left            =   525
            TabIndex        =   68
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   635
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
         Begin CisText_V60.CisText iB3_Day 
            Height          =   360
            Index           =   2
            Left            =   930
            TabIndex        =   69
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   635
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
         Begin CisText_V60.CisText iB3_Day 
            Height          =   360
            Index           =   3
            Left            =   1335
            TabIndex        =   70
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   635
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
         Begin CisText_V60.CisText iB3_Day 
            Height          =   360
            Index           =   4
            Left            =   1740
            TabIndex        =   71
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   635
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
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   27
         Left            =   1515
         Top             =   2370
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1349
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisText_V60.CisText iB3_Keisu3 
            Height          =   360
            Left            =   60
            TabIndex        =   58
            Top             =   360
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   635
            cFormat         =   "0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0.00"
            MaxLength       =   4
            cDataType       =   2
            cDataReplace    =   1
            cFbComma        =   0
            cFcDecimal      =   2
            cFdAutoFormat   =   1
            cGFormat        =   "0.00"
            cILength        =   1
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   28
         Left            =   2235
         Top             =   2370
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "LT"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB3_LT 
            Height          =   360
            Left            =   90
            TabIndex        =   59
            Top             =   360
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
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
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   26
         Left            =   795
         Top             =   2370
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1349
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisText_V60.CisText iB3_Keisu2 
            Height          =   360
            Left            =   60
            TabIndex        =   57
            Top             =   360
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   635
            cFormat         =   "0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0.00"
            MaxLength       =   4
            cDataType       =   2
            cDataReplace    =   1
            cFbComma        =   0
            cFcDecimal      =   2
            cFdAutoFormat   =   1
            cGFormat        =   "0.00"
            cILength        =   1
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   25
         Left            =   75
         Top             =   2370
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1349
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisText_V60.CisText iB3_Keisu1 
            Height          =   360
            Left            =   60
            TabIndex        =   56
            Top             =   360
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   635
            cFormat         =   "0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0.00"
            MaxLength       =   4
            cDataType       =   2
            cDataReplace    =   1
            cFbComma        =   0
            cFcDecimal      =   2
            cFdAutoFormat   =   1
            cGFormat        =   "0.00"
            cILength        =   1
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   30
         Left            =   4965
         Top             =   2370
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "ój ì˙ éw íË"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.CheckBox B3Chk_Week 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ã‡"
            Height          =   285
            Index           =   4
            Left            =   2610
            TabIndex        =   66
            Top             =   390
            Width           =   525
         End
         Begin VB.CheckBox B3Chk_Week 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ñÿ"
            Height          =   285
            Index           =   3
            Left            =   1980
            TabIndex        =   65
            Top             =   390
            Width           =   525
         End
         Begin VB.CheckBox B3Chk_Week 
            BackColor       =   &H00C0C0C0&
            Caption         =   "êÖ"
            Height          =   285
            Index           =   2
            Left            =   1380
            TabIndex        =   64
            Top             =   390
            Width           =   525
         End
         Begin VB.CheckBox B3Chk_Week 
            BackColor       =   &H00C0C0C0&
            Caption         =   "âŒ"
            Height          =   285
            Index           =   1
            Left            =   810
            TabIndex        =   63
            Top             =   390
            Width           =   525
         End
         Begin VB.CheckBox B3Chk_Week 
            BackColor       =   &H00C0C0C0&
            Caption         =   "åé"
            Height          =   285
            Index           =   0
            Left            =   210
            TabIndex        =   62
            Top             =   390
            Width           =   525
         End
         Begin Cis3D_v60.CIS3D B3lb_YoubiHC 
            Height          =   285
            Left            =   2490
            Top             =   45
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            BackColor       =   16777152
            Caption         =   "î≠íçóL"
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
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D Back_Juchu 
         Height          =   1935
         Index           =   0
         Left            =   75
         Top             =   2370
         Width           =   10440
         _ExtentX        =   18415
         _ExtentY        =   3413
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
         cAlingnment     =   6
         cPositionY      =   30
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   2
         Left            =   6600
         Top             =   840
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   1349
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisYMD_With_Btn.CisYMDwB iB3_YSYmd 
            Height          =   330
            Left            =   90
            TabIndex        =   42
            Top             =   360
            Width           =   1695
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
            cChkResult      =   0   'False
            cCalenderPosition=   1
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
         Begin CisYMD_With_Btn.CisYMDwB iB3_YEYmd 
            Height          =   330
            Left            =   2100
            TabIndex        =   43
            Top             =   360
            Width           =   1695
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
            cChkResult      =   0   'False
            cCalenderPosition=   1
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
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Å`"
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   1830
            TabIndex        =   99
            Top             =   420
            Width           =   435
         End
      End
      Begin Cis3D_v60.CIS3D Back_Juchu 
         Height          =   765
         Index           =   13
         Left            =   75
         Top             =   1605
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "ì`ï[éÌóﬁ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB3_DenSyuJ 
            Height          =   345
            Left            =   90
            TabIndex        =   100
            Top             =   360
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
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
         Begin Cis3D_v60.CIS3D B3lb_DenSyuJ 
            Height          =   330
            Left            =   330
            Top             =   360
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   582
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D Back_Juchu 
         Height          =   765
         Index           =   1
         Left            =   7245
         Top             =   1605
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   1349
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
         cAlingnment     =   6
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D Back_Juchu 
            Height          =   825
            Index           =   15
            Left            =   0
            Top             =   0
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   1455
            ForeColor       =   16711680
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
            cAlingnment     =   6
            Begin Cis3D_v60.CIS3D B3lb_SyukaKNM 
               Height          =   345
               Left            =   1035
               Top             =   360
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   609
               BackColor       =   16777152
               Caption         =   "XXXXXXXXX1XXXXX"
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
               cFont3D         =   2
               cAlingnment     =   1
               cPositionX      =   30
            End
            Begin CisText_V60.CisText iB3_SyukaK 
               Height          =   330
               Left            =   45
               TabIndex        =   106
               Top             =   360
               Width           =   975
               _ExtentX        =   1720
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
               Text            =   "XXXXXXX"
               MaxLength       =   7
            End
         End
      End
      Begin Cis3D_v60.CIS3D Back_HachuAdrs 
         Height          =   765
         Left            =   4965
         Top             =   840
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "ÉAÉhÉåÉX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.CheckBox B3Chk_Auto 
            BackColor       =   &H00C0C0C0&
            Caption         =   "éËìÆ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1065
            TabIndex        =   41
            Top             =   330
            Width           =   510
         End
         Begin CisText_V60.CisText iB3_Adress 
            Height          =   360
            Left            =   60
            TabIndex        =   40
            Top             =   345
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   635
            cFormat         =   "0000000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0000000"
            MaxLength       =   7
            cDataType       =   1
            cDataReplace    =   1
            cFaZero         =   0
            cFbComma        =   0
            cFdAutoFormat   =   1
            cGFormat        =   "#######"
            cILength        =   7
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   765
         Index           =   22
         Left            =   9765
         Top             =   75
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "ç≈èIáÇ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D B3lb_HNo 
            Height          =   330
            Left            =   120
            Top             =   360
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   765
         Left            =   3000
         Top             =   75
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1349
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
         cPositionY      =   20
         Begin CisText_V60.CisText iB3_Cycle1 
            Height          =   360
            Left            =   15
            TabIndex        =   29
            Top             =   360
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   635
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
            Text            =   "#"
            MaxLength       =   1
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "#"
            cILength        =   1
         End
         Begin CisText_V60.CisText iB3_Cycle2 
            Height          =   360
            Left            =   255
            TabIndex        =   30
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   635
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
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
         Begin CisText_V60.CisText iB3_Cycle3 
            Height          =   360
            Left            =   615
            TabIndex        =   31
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   635
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
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D40 
         Height          =   765
         Left            =   8340
         Top             =   75
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "áA"
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
         cPositionX      =   60
         cPositionY      =   60
         Begin CisText_V60.CisText iB3_Syuyo2 
            Height          =   360
            Left            =   390
            TabIndex        =   36
            Top             =   30
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
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
         Begin CisText_V60.CisText iB3_Youki2 
            Height          =   360
            Left            =   60
            TabIndex        =   37
            Top             =   390
            Width           =   1335
            _ExtentX        =   2355
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
            Text            =   "XXXXXXXXXX"
            MaxLength       =   10
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D10 
         Height          =   765
         Left            =   4020
         Top             =   75
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "á@"
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
         cPositionX      =   900
         cPositionY      =   60
         Begin CisText_V60.CisText iB3_Syuyo1 
            Height          =   360
            Left            =   1155
            TabIndex        =   32
            Top             =   30
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
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
         Begin CisText_V60.CisText iB3_Youki1 
            Height          =   360
            Left            =   810
            TabIndex        =   33
            Top             =   390
            Width           =   1335
            _ExtentX        =   2355
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
            Text            =   "XXXXXXXXXX"
            MaxLength       =   10
            IMEMode         =   2
         End
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   345
            Index           =   16
            Left            =   60
            Top             =   360
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   609
            ForeColor       =   16711680
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
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   2
            cPositionY      =   30
         End
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   345
            Index           =   15
            Left            =   60
            Top             =   30
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   609
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
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   2
            cPositionY      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   765
         Index           =   17
         Left            =   1260
         Top             =   75
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "SKãÊï™"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.OptionButton B3Op_SK 
            BackColor       =   &H00C0C0C0&
            Caption         =   "éwé¶"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   1050
            TabIndex        =   28
            Top             =   375
            Width           =   660
         End
         Begin VB.OptionButton B3Op_SK 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ç©ÇÒÇŒÇÒ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   30
            TabIndex        =   27
            Top             =   375
            Value           =   -1  'True
            Width           =   1035
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   765
         Left            =   1515
         Top             =   840
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "å_Å@ñÒÅ@êÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB3_KTorcd 
            Height          =   360
            Left            =   60
            TabIndex        =   39
            Top             =   330
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
         Begin Cis3D_v60.CIS3D B3lb_KTorcd 
            Height          =   360
            Left            =   1050
            Top             =   330
            Width           =   2295
            _ExtentX        =   4048
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
      Begin Cis3D_v60.CIS3D CIS3D20 
         Height          =   765
         Left            =   6165
         Top             =   75
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   " ïœ  çX  ì˙ - ï÷"
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
         Begin CisYMD_With_Btn.CisYMDwB iB3_HYmd 
            Height          =   330
            Left            =   60
            TabIndex        =   34
            Top             =   390
            Width           =   1695
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
         Begin CisText_V60.CisText iB3_HBin 
            Height          =   360
            Left            =   1770
            TabIndex        =   35
            Top             =   390
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
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
      Begin Cis3D_v60.CIS3D Back_Seban 
         Height          =   765
         Left            =   75
         Top             =   75
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1349
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
         Begin CisText_V60.CisText iB3_Seban 
            Height          =   360
            Left            =   45
            TabIndex        =   26
            Top             =   360
            Width           =   1110
            _ExtentX        =   1958
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
            Text            =   "XXXXXXXX"
            MaxLength       =   8
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   765
         Index           =   18
         Left            =   75
         Top             =   840
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "íuÅ@èÍ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB3_Okiba 
            Height          =   360
            Left            =   60
            TabIndex        =   38
            Top             =   330
            Width           =   1335
            _ExtentX        =   2355
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
            Text            =   "XXXXXXXXXX"
            MaxLength       =   10
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D Back_Juchu 
         Height          =   765
         Index           =   7
         Left            =   6600
         Top             =   840
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "îıÅ@Å@Å@Å@Å@Å@çl"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB3_JBiko 
            Height          =   360
            Left            =   75
            TabIndex        =   98
            Top             =   330
            Width           =   3780
            _ExtentX        =   6668
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   30
            IMEMode         =   4
         End
      End
      Begin Cis3D_v60.CIS3D Back_Juchu 
         Height          =   765
         Index           =   12
         Left            =   4710
         Top             =   1605
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "éÛóÃèëî≠çs"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.CheckBox B3Chk_JuryoJ 
            BackColor       =   &H00C0C0C0&
            Caption         =   "éÛíçéû"
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
            Left            =   105
            TabIndex        =   104
            Top             =   420
            Width           =   1080
         End
      End
      Begin Cis3D_v60.CIS3D Back_Juchu 
         Height          =   765
         Index           =   9
         Left            =   3645
         Top             =   1605
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "éÛóÃèë"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.CheckBox B3Chk_PrtkbJJ 
            BackColor       =   &H00C0C0C0&
            Caption         =   "î≠çsóL"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   90
            TabIndex        =   103
            Top             =   405
            Width           =   960
         End
      End
      Begin Cis3D_v60.CIS3D Back_Juchu 
         Height          =   765
         Index           =   10
         Left            =   1515
         Top             =   1605
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "î[ïièë"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.CheckBox B3Chk_PrtkbNJ 
            BackColor       =   &H00C0C0C0&
            Caption         =   "î≠çsóL"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   60
            TabIndex        =   101
            Top             =   405
            Width           =   945
         End
      End
      Begin Cis3D_v60.CIS3D Back_Juchu 
         Height          =   765
         Index           =   17
         Left            =   2580
         Top             =   1605
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "éxããèë"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.CheckBox B3Chk_PrtkbSJ 
            BackColor       =   &H00C0C0C0&
            Caption         =   "î≠çsóL"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   75
            TabIndex        =   102
            Top             =   405
            Width           =   960
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   10
         Left            =   1515
         Top             =   1605
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   1349
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
         cAlingnment     =   6
         cPositionY      =   30
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   9
         Left            =   1515
         Top             =   1605
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "Å@Å@ë„çså_ñÒêÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.CheckBox ChkB3_Daiko 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Åyë„çsÅz"
            Height          =   315
            Left            =   60
            TabIndex        =   109
            Top             =   30
            Width           =   1290
         End
         Begin CisText_V60.CisText iB3_DaikoK 
            Height          =   360
            Left            =   60
            TabIndex        =   110
            Top             =   360
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
         Begin Cis3D_v60.CIS3D B3lb_DaikoK 
            Height          =   360
            Left            =   1050
            Top             =   360
            Width           =   1890
            _ExtentX        =   3334
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
      Begin Cis3D_v60.CIS3D Back_Juchu 
         Height          =   765
         Index           =   32
         Left            =   7245
         Top             =   1605
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "é©é–çHãÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.ComboBox B3Comb_Jisya 
            BackColor       =   &H00C0E0FF&
            Height          =   360
            Left            =   60
            Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
            TabIndex        =   107
            Top             =   330
            Width           =   1335
         End
      End
      Begin Cis3D_v60.CIS3D Back_Juchu 
         Height          =   765
         Index           =   21
         Left            =   5955
         Top             =   1605
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "éwé¶èëî≠çs"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.CheckBox B3Chk_Sijisyo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "î≠çsóL"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   165
            TabIndex        =   105
            Top             =   390
            Width           =   960
         End
         Begin VB.CheckBox B3Chk_GouhoJ 
            BackColor       =   &H00C0C0C0&
            Caption         =   "çÜå˚"
            Height          =   270
            Left            =   255
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   405
            Width           =   930
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   765
         Index           =   22
         Left            =   1515
         Top             =   1605
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "î[ïièëî≠çs"
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
         Begin VB.CheckBox B3Chk_PrtkbSN 
            BackColor       =   &H00C0C0C0&
            Caption         =   "î≠çsóL"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1275
            TabIndex        =   47
            Top             =   405
            Width           =   945
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hachu 
         Height          =   750
         Index           =   23
         Left            =   75
         Top             =   3135
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   1323
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
      End
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   510
         Left            =   75
         Top             =   4305
         Width           =   10440
         _ExtentX        =   18415
         _ExtentY        =   900
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
         cAlingnment     =   6
         cPositionY      =   30
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13890
      Top             =   1140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   120
      ScaleHeight     =   885
      ScaleWidth      =   7245
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   510
      Width           =   7305
      Begin Cis3D_v60.CIS3D BLB_HINBN 
         Height          =   795
         Left            =   60
         Top             =   60
         Width           =   3270
         _ExtentX        =   5768
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
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   330
            Left            =   60
            TabIndex        =   0
            Top             =   360
            Width           =   3210
            _ExtentX        =   5662
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
            IMEMode         =   2
            cAutoEject      =   0   'False
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   795
         Index           =   14
         Left            =   3330
         Top             =   60
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "ï\ é¶ ïi î‘"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D B1lb_EHinbn 
            Height          =   345
            Left            =   60
            Top             =   360
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   609
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      End
   End
   Begin VB.Timer Timer3 
      Left            =   1980
      Top             =   810
   End
   Begin VB.PictureBox B2_Area1 
      Height          =   765
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   10605
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   5370
      Width           =   10665
      Begin VB.PictureBox B2_Area2 
         Height          =   645
         Left            =   30
         ScaleHeight     =   585
         ScaleWidth      =   7140
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   60
         Width           =   7200
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   480
            Index           =   11
            Left            =   4935
            Top             =   60
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   847
            ForeColor       =   16711680
            Caption         =   "ïiñ⁄"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            cPositionX      =   60
            Begin VB.OptionButton B2Op_Hinmoku 
               BackColor       =   &H00C0FFC0&
               Caption         =   "î≠íç"
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   1
               Left            =   1380
               TabIndex        =   25
               Top             =   75
               Value           =   -1  'True
               Width           =   720
            End
            Begin VB.OptionButton B2Op_Hinmoku 
               BackColor       =   &H00C0C0C0&
               Caption         =   "éÛíç"
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   0
               Left            =   600
               TabIndex        =   24
               Top             =   75
               Width           =   735
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   480
            Left            =   60
            Top             =   60
            Width           =   3900
            _ExtentX        =   6879
            _ExtentY        =   847
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
            cAlingnment     =   1
            cPositionX      =   60
            Begin CisText_V60.CisText iB2_Torcd 
               Height          =   360
               Left            =   825
               TabIndex        =   22
               Top             =   75
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
               IMEMode         =   2
            End
            Begin Cis3D_v60.CIS3D B2lb_Torcd 
               Height          =   345
               Left            =   1815
               Top             =   75
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   609
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
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   480
            Index           =   9
            Left            =   3960
            Top             =   60
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   847
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
            cAlingnment     =   1
            cPositionX      =   60
            Begin CisText_V60.CisText iB2_Ukeir 
               Height          =   360
               Left            =   555
               TabIndex        =   23
               Top             =   60
               Width           =   345
               _ExtentX        =   609
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
      End
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   1620
      Left            =   120
      ScaleHeight     =   1560
      ScaleWidth      =   15045
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   1500
      Width           =   15105
      Begin Cis3D_v60.CIS3D CIS3D16 
         Height          =   735
         Left            =   9240
         Top             =   795
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "ê∂ãÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_Abc 
            Height          =   360
            Left            =   150
            TabIndex        =   17
            Top             =   330
            Width           =   255
            _ExtentX        =   450
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
            Text            =   "X"
            MaxLength       =   1
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D12 
         Height          =   735
         Left            =   9780
         Top             =   795
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "ë„ï\ìæà”êÊ"
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
         cPositionY      =   60
         Begin CisText_V60.CisText iB1_DaiCD 
            Height          =   375
            Left            =   60
            TabIndex        =   18
            Top             =   330
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
            IMEMode         =   2
         End
         Begin Cis3D_v60.CIS3D B1lb_DaiNM 
            Height          =   360
            Left            =   1035
            Top             =   330
            Width           =   1710
            _ExtentX        =   3016
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
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   735
         Left            =   13245
         Top             =   60
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "ïiãÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D B1lb_Hinku 
            Height          =   345
            Left            =   285
            Top             =   345
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   609
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
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin CisText_V60.CisText iB1_Hinku 
            Height          =   360
            Left            =   30
            TabIndex        =   11
            Top             =   345
            Width           =   255
            _ExtentX        =   450
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
            Text            =   "X"
            MaxLength       =   1
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D91 
         Height          =   735
         Left            =   14175
         Top             =   60
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "é¿êî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.CheckBox B1Chk_JSKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ç»Çµ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   45
            TabIndex        =   12
            Top             =   330
            Width           =   720
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   735
         Index           =   35
         Left            =   5400
         Top             =   795
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "ÉCÉÅÅ[ÉWÉtÉ@ÉCÉãñº"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D PB_Image 
            Height          =   375
            Left            =   3420
            Top             =   330
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   661
            Caption         =   "Åd"
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
            cButton         =   -1  'True
         End
         Begin CisText_V60.CisText iB1_Image 
            Height          =   360
            Left            =   90
            TabIndex        =   16
            Top             =   330
            Width           =   3375
            _ExtentX        =   5953
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
            Text            =   $"CXM0030.frx":0000
            MaxLength       =   255
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   735
         Left            =   3000
         Top             =   795
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "å^ èä ç› ín"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_KataTor 
            Height          =   360
            Left            =   60
            TabIndex        =   15
            Top             =   345
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
         Begin Cis3D_v60.CIS3D B1lb_KataTor 
            Height          =   360
            Left            =   1050
            Top             =   330
            Width           =   1230
            _ExtentX        =   2170
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
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   735
         Left            =   1200
         Top             =   795
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "å^áÇ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_KataNo 
            Height          =   360
            Left            =   45
            TabIndex        =   14
            Top             =   330
            Width           =   1725
            _ExtentX        =   3043
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
            Text            =   "XXXXXXXXXXXXX"
            MaxLength       =   13
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   735
         Index           =   21
         Left            =   12585
         Top             =   795
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "çﬁóø"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D B1lb_Zairyo 
            Height          =   360
            Left            =   60
            Top             =   330
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "ZZZ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   735
         Index           =   10
         Left            =   45
         Top             =   795
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "íS ìñ é“"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_Tanto 
            Height          =   360
            Left            =   30
            TabIndex        =   13
            Top             =   330
            Width           =   495
            _ExtentX        =   873
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
            Text            =   "XXX"
            MaxLength       =   3
         End
         Begin Cis3D_v60.CIS3D B1lb_Tanto 
            Height          =   345
            Left            =   540
            Top             =   330
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   609
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   735
         Index           =   8
         Left            =   990
         Top             =   60
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "é–ì‡îwî‘"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_SSeban 
            Height          =   360
            Left            =   180
            TabIndex        =   2
            Top             =   330
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
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   735
         Index           =   4
         Left            =   60
         Top             =   60
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "é‘ å^"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_Syakei 
            Height          =   360
            Left            =   120
            TabIndex        =   1
            Top             =   330
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
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   735
         Left            =   7200
         Top             =   60
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "äÓñ{é˚óeêî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_KSyuyo 
            Height          =   360
            Left            =   180
            TabIndex        =   4
            Top             =   330
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   635
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
      End
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   735
         Left            =   8520
         Top             =   60
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "äÓñ{é˚óeäÌ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_KUtsuwa 
            Height          =   360
            Left            =   45
            TabIndex        =   5
            Top             =   330
            Width           =   1365
            _ExtentX        =   2408
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
            Text            =   "XXXXXXXXXX"
            MaxLength       =   10
            IMEMode         =   2
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   735
         Index           =   12
         Left            =   2145
         Top             =   60
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "ïi        ñº"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_HinNM 
            Height          =   360
            Left            =   60
            TabIndex        =   3
            Top             =   330
            Width           =   5025
            _ExtentX        =   8864
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   40
            IMEMode         =   4
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   735
         Index           =   0
         Left            =   9990
         Top             =   60
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "ïiî‘ï“èW"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_Edit 
            Height          =   360
            Left            =   180
            TabIndex        =   6
            Top             =   330
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
            cDataReplace    =   1
            cFaZero         =   0
            cFbComma        =   0
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   735
         Index           =   1
         Left            =   13155
         Top             =   795
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "ïîïiãÊï™"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.OptionButton B1Op_Buhin 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ÉKÉX"
            Height          =   315
            Index           =   1
            Left            =   990
            TabIndex        =   20
            Top             =   390
            Width           =   795
         End
         Begin VB.OptionButton B1Op_Buhin 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ïîïi"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   390
            Value           =   -1  'True
            Width           =   795
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9x 
         Height          =   735
         Left            =   11055
         Top             =   60
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "ç› å… î≠ íç"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.ComboBox B1Comb_Zaiko 
            BackColor       =   &H00C0E0FF&
            Height          =   360
            Left            =   15
            Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
            TabIndex        =   7
            Top             =   330
            Width           =   1155
         End
         Begin CisText_V60.CisText iB1_ZaikoDay 
            Height          =   360
            Left            =   1185
            TabIndex        =   8
            Top             =   345
            Visible         =   0   'False
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   635
            cFormat         =   "#0.0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "#0.0"
            MaxLength       =   4
            cDataType       =   2
            cDataReplace    =   1
            cFbComma        =   0
            cFcDecimal      =   1
            cFdAutoFormat   =   1
            cGFormat        =   "#0.0"
            cILength        =   2
         End
         Begin CisText_V60.CisText iB1_Zaiko 
            Height          =   360
            Left            =   1185
            TabIndex        =   9
            Top             =   360
            Visible         =   0   'False
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   635
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
            cFbComma        =   0
            cFdAutoFormat   =   1
            cGFormat        =   "###"
            cILength        =   3
         End
         Begin CisText_V60.CisText iB1_ZaikoL 
            Height          =   375
            Left            =   1170
            TabIndex        =   10
            Top             =   330
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
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
         Begin VB.Label Zaiko_Label 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "ì˙"
            Height          =   240
            Left            =   1890
            TabIndex        =   94
            Top             =   450
            Visible         =   0   'False
            Width           =   240
         End
      End
   End
   Begin VB.PictureBox B1_Area2 
      Height          =   2130
      Left            =   120
      ScaleHeight     =   2070
      ScaleWidth      =   15045
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   3180
      Width           =   15105
      Begin vsFlexLib.vsFlexArray vsGrid1 
         Height          =   1980
         Left            =   60
         TabIndex        =   21
         Top             =   45
         Width           =   14925
         _Version        =   196608
         _ExtentX        =   26326
         _ExtentY        =   3492
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   $"CXM0030.frx":0103
         Rows            =   5
         Cols            =   47
         BackColor       =   12648447
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   330
         SelectionMode   =   1
         BackColorAlternate=   12648447
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
      Left            =   7470
      ScaleHeight     =   660
      ScaleWidth      =   7725
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   510
      Width           =   7785
      Begin CisBtn_60.CisBtn PB_Zairyo 
         Height          =   585
         Left            =   45
         Top             =   30
         Visible         =   0   'False
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
         BtnCaption      =   "ä÷òA"
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
         PFBackColor     =   8388736
      End
      Begin CisBtn_60.CisBtn PB_ImageDisp 
         Height          =   585
         Left            =   870
         Top             =   30
         Visible         =   0   'False
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
         BtnCaption      =   "≤“∞ºﬁ"
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
         PFForeColor     =   12648447
         PFShadowColor   =   0
         PFBackColor     =   192
      End
      Begin CisBtn_60.CisBtn PB_Chg 
         Height          =   585
         Left            =   4290
         Tag             =   "TOR"
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
         BtnCaption      =   "éÊà¯êÊ"
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
         PFCaption       =   "( F9 )"
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   5145
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
         Left            =   6000
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
         Left            =   6855
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
         Left            =   1725
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
         Left            =   2580
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
         Left            =   4290
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
         Left            =   3435
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
      Left            =   990
      TabIndex        =   85
      Text            =   "Text1"
      Top             =   540
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Left            =   930
      Top             =   660
   End
   Begin VB.Timer Timer2 
      Left            =   1440
      Top             =   720
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
      Caption         =   "Åy ïiî‘É}ÉXÉ^ÉÅÉìÉeÉiÉìÉX  Åz"
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
         TabIndex        =   93
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
         Left            =   14400
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
   Begin VB.PictureBox B3_Area2 
      Height          =   4890
      Left            =   10830
      ScaleHeight     =   4920
      ScaleMode       =   0  '’∞ªﬁ∞
      ScaleWidth      =   4380
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   6180
      Width           =   4440
      Begin Cis3D_v60.CIS3D CIS3D31 
         Height          =   4742
         Left            =   30
         Top             =   60
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   8361
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
         Begin vsFlexLib.vsFlexArray vsGrid2 
            Height          =   4140
            Left            =   75
            TabIndex        =   112
            TabStop         =   0   'False
            Top             =   555
            Width           =   4200
            _Version        =   196608
            _ExtentX        =   7408
            _ExtentY        =   7302
            _StockProps     =   228
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "çHèá|ã@äB(çÏã∆)|äOíçâ¡çH   |ã@äBáÇ      |€ŒﬁØƒ|çHíˆ         |äÓèÄçHêî1|äÓèÄçHêî2"
            Rows            =   12
            BackColor       =   12648384
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   330
            SelectionMode   =   1
            BackColorAlternate=   12648384
         End
         Begin Cis3D_v60.CIS3D CIS3D34 
            Height          =   405
            Left            =   90
            Top             =   75
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   714
            ForeColor       =   16711680
            Caption         =   "äÓèÄéûä‘åv"
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
            cPositionX      =   300
            Begin Cis3D_v60.CIS3D B3lb_KosuKei 
               Height          =   285
               Left            =   2400
               Top             =   60
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "##0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            End
         End
      End
   End
   Begin Cis3D_v60.CIS3D K_Sykbnm2 
      Height          =   765
      Left            =   10860
      Top             =   5370
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1349
      Caption         =   "í« â¡"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   15.75
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
   Begin Cis3D_v60.CIS3D PB_Kotei 
      Height          =   660
      Left            =   12600
      Top             =   5475
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1164
      BackColor       =   8454143
      Caption         =   "çH èá ä« óù(F11)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   3
      cAlingnment     =   6
      cButton         =   -1  'True
      cPositionY      =   65
   End
End
Attribute VB_Name = "CXM0030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :   ïiî‘É}ÉXÉ^ÉÅÉìÉeÉiÉìÉX
'**       ÉtÉHÅ[ÉÄID    :   CXM0030
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :   2003/11/25  By CIS SIMPR-A
'**       ïœ  çX  ì˙    :   2004/04/05  ïsãÔçáèCê≥(ñﬂÇÈÉ{É^Éì)
'**       ïœ  çX  ì˙    :   2004/04/08  ïsãÔçáèCê≥(éÛíçéûÇÃéÛóÃèëÉRÉìÉ{É{ÉbÉNÉX)
'**       ïœ  çX  ì˙    :   2004/04/15  ïsãÔçáèCê≥(éÊà¯êÊñºï\é¶ÅEABCãÊï™ï\é¶)
'**       ïœ  çX  ì˙    :   2004/05/28  é–ì‡îwî‘çÜÉJÉiì¸óÕ
'**                                     é–ì‡îwî‘çÜÅEîwî‘çÜîºäpÉ`ÉFÉbÉN
'**       ïœ  çX  ì˙    :   2004/05/31  ïiî‘åüçıÇïiî‘éÊà¯êÊåüçıÇ÷ïœçX
'**       ïœ  çX  ì˙    :   2004/06/28  ïsãÔçáèCê≥ÅiÅj
'**       ïœ  çX  ì˙    :   2004/06/29  éÊà¯êÊÇ™é–ì‡éËîzãÊï™Çé–ì‡Ç…ê›íËÇ∑ÇÈÇ∆èCê≥ïsâ¬Ç…Ç»ÇÈà»äOÇÃéûÅAéËîzãÊï™Ç™é–ì‡ÇÕÉGÉâÅ[Ç∆Ç∑ÇÈ
'**       ïœ  çX  ì˙    :   2004/09/24  ïsãÔçáèCê≥ÅiÉTÉCÉNÉãÇÉNÉäÉAéûÅAçXêVÇ≥ÇÍÇ»Ç¢Åj
'**       ïœ  çX  ì˙    :   2004/12/13  ã@äBáÇèÓïÒí«â¡
'**       ïœ  çX  ì˙    :   2005/01/31  By CIS î≠íçÇÃéûÅAì`ï[éÌóﬁÇ…é©ããÉ`ÉPÉbÉgÅEéxããÉ`ÉPÉbÉgÇí«â¡
'**       ïœ  çX  ì˙    :   2005/09/30  By CIS PNOÇÃåüçıí«â¡
'**       ïœ  çX  ì˙    :   2006/02/02  By CIS - INIÉtÉ@ÉCÉã-ÉNÉâÉCÉAÉìÉgëŒâûÅiTEXTèoóÕêÊÅj
'**       ïœ  çX  ì˙    :   2006/02/16  By CIS ±ƒﬁ⁄Ωèdï°¡™Ø∏ïœçXÅiïiñ⁄íPà Ç…ïœçXÅj
'**                                             É|ÉPÉbÉgç≈ëÂêîê›íËíl(ñºèÃÉ}ÉXÉ^)îΩâf
'**       ïœ  çX  ì˙    :   2006/02/23  BY CIS ïiî‘éÊà¯êÊèÓïÒí«â¡éûÅAéÊà¯êÊÇ™å_ñÒêÊÇÃéûÅAå_ñÒêÊÇÃï\é¶Ç™Ç≥ÇÍÇ»Ç¢
'**       ïœ  çX  ì˙    :   2006/03/16  By CIS ç\ê¨œΩ¿“›√≈›ΩòAågéû,√ﬁ∞¿∞Õﬁ∞ΩÇÃçƒµ∞Ãﬂ›ÇçsÇ¡ÇƒÇÈÇÃÇèCê≥
'**       ïœ  çX  ì˙    :   2006/03/28  By CIS çÌèúèàóùéûÅAÇ©ÇÒÇŒÇÒó¨ìÆä«óùÉeÅ[ÉuÉãçÌèúèàóùí«â¡
'**       ïœ  çX  ì˙    :   2006/12/19  By CIS ÉCÉÅÅ[ÉWÉ{É^ÉìÇâüâ∫éûÇÃèâä˙à íuÇ™ä¬ã´ê›íËÇÃà íuÇ∆à·Ç§ÇÃÇèCê≥
'**       ïœ  çX  ì˙    :   2007/12/10  By CIS î[ïièëî≠çsãÊï™•éxããèëî≠çsãÊï™•éÛóÃèëî≠çsãÊï™Å@ÇÃí«â¡
'**                                             ÅEî[ïièëå`ë‘•éÛóÃèëå`ë‘Çï\é¶çÌèú(B3Comb_Nouhin,B3Comb_Jyuryo,B3Comb_NouhinJ,B3Comb_JyuryoJ)
'**                                             ÅEñºèÃÉ}ÉXÉ^"ì`ï[éÌóﬁéÛíç•î≠íç"ÇÃâpêîéö1Ç…î[ïièëå`ë‘•éÛóÃèëå`ë‘Çê›íË
'**                                             Å@Åiî[ïièëå`ë‘Ç∆éÛóÃèëå`ë‘ÇÕìØÇ∂Ç∆Ç∑ÇÈÅBäÓñ{ÇÕé©ããÇÕî≠çsÅAéxããÇÕÃß≤ŸèoóÕÇ∆Ç∑ÇÈÅj
'**       ïœ  çX  ì˙    :   2007/12/14  By CIS éËîzãÊï™=éÛããÇÃéûÅAë„çsãÊï™ÅEë„çså_ñÒêÊÅië„çsêÊÇÃå_ñÒêÊÅjÇÃí«â¡
'**       ïœ  çX  ì˙    :   2008/02/04  By CIS SIMPR-AïWèÄÇÊÇËà⁄çs
'**                                            êÁë„ìcçHã∆ÉJÉXÉ^É}ÉCÉY
'**       ïœ  çX  ì˙    :   2008/03/31  By CIS ç\ê¨É}ÉXÉ^ÇÊÇËà⁄çsÇµÇƒÇ´ÇΩèÍçáÇ…ïiî‘ÇÃÉÅÉìÉeÉiÉìÉXëSëÃÇçsÇ¶ÇÈÇÊÇ§ïœçX
'**       ïœ  çX  ì˙    :   2008/06/09  By CIS éËîzãÊï™ = éxãã í«â¡,
'**       ïœ  çX  ì˙    :   2008/06/12  By CIS çwì¸íPâøÉ}ÉXÉ^Ç÷ÇÃà⁄çséûèCê≥
'**       ïœ  çX  ì˙    :   2009/07/06  By CIS êVe-Ç©ÇÒÇŒÇÒëŒâû èoâ◊èÍí«â¡
'**       ïœ  çX  ì˙    :   2009/07/22  By CIS ìyäÚí√Ver•••çHèáä«óùí«â¡
'**       ïœ  çX  ì˙    :   2009/10/16  By CIS ìyäÚí√Ver
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' Ã´∞—±∏√®Ãﬁ Ã◊∏ﬁ
    Dim ProcHB          As String * 2   ' ì¸óÕàÊêÿë÷
    Dim SyoriKB         As String       ' èàóùãÊï™(êVãK•èCê≥•çÌèú ìô)
    Dim SyoriNM         As String       ' èàóùñºèÃ(êVãK•èCê≥•çÌèú ìô)
    Dim SyoriKB2        As String       ' èàóùãÊï™(êVãK•èCê≥•çÌèú ìô)
    Dim SyoriNM2        As String       ' èàóùñºèÃ(êVãK•èCê≥•çÌèú ìô)
    Dim mCHK            As Boolean
    Dim RowColFlg       As Boolean
    Dim GridBack        As String
    
    Dim DayString()     As String
    
    Dim mGotColor       As Long
    Dim mLostColor      As Long
    
    Dim CisVsGrid       As New CisVsGrid3
    'Dim CisVsGridS      As New CisVsGrid3          2009/07/22
    '+------------------------------------------- insert 2004/12/13
    'Dim CisVsGridK      As New CisVsGrid3          2009/07/22
    '+------------------------------------------- insert 2004/12/13 End
    
    '============================================*** 2009/07/22 Start
    Dim SyoriKB_K       As String       ' èàóùãÊï™(êVãK•èCê≥•çÌèú ìô)
    Dim SyoriNM_K       As String       ' èàóùñºèÃ(êVãK•èCê≥•çÌèú ìô)
    Dim K_Select_Data   As Boolean      ' ÉJÅ[É\Éãà íu
    Dim CisVsGrid2      As New CisVsGrid3
'---------------------------------------------------' 2009.10.20 del start
''''    Dim CisVsGrid3      As New CisVsGrid3
''''    Dim CisVsGrid4      As New CisVsGrid3
'---------------------------------------------------' 2009.10.20 del end
''''''''    Dim CisComboKZaiko  As CisCombo
''''''''    Dim CisComboKMZaiko As CisCombo
    Dim CisComboTest As CisCombo
    '============================================*** 2009/07/22 End
    
    Dim CisComboTehai   As CisCombo
    Dim CisComboZaiko   As CisCombo
'''2007/12/11DEL    Dim CisComboHNou    As CisCombo
'''2007/12/11DEL    Dim CisComboHJyu    As CisCombo
    Dim CisComboHMei    As CisCombo
'''2007/12/11DEL    Dim CisComboJNou    As CisCombo
'''2007/12/11DEL    Dim CisComboJJyu    As CisCombo
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
    Dim CisComboJsya    As CisCombo
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
        
    Dim mIx             As Long
    Dim mIx2            As Long
    Dim mIx3            As Long
    Dim mStr            As String
    
    Dim ActObj          As Object
    Dim ActRow          As Integer
    
    Dim KCol            As Long         ' îÒï\é¶çÄñ⁄äJén∂◊—
    Dim SCol()          As Long         ' ï\é¶çÄñ⁄∂◊—

    Dim mTorFlg         As Boolean
    
    Dim mBodyErrFld     As Object
    
    Dim mImage          As String
    Dim maxPNo          As Byte         ' ç≈ëÂÉ|ÉPÉbÉgáÇ ( 2006/02/16 ADD )

'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2006/12/19 Update Satart
    Dim BD1Chk_Flg      As Byte
'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2006/12/19 Update End
'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2007/12/11 Update Satart
    Dim SaveKeitai      As String
'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2007/12/11 Update End

'-------------------------------------------------------------------------------' 2009.10.16 add start
'    Dim SAVE_Select     As String
    Dim PosiSave0X      As Integer
    Dim PosiSave0Y      As Integer
    Dim PosiSave1X      As Integer
    Dim PosiSave1Y      As Integer
    Dim PosiSave1W      As Integer
'-------------------------------------------------------------------------------' 2009.10.16 add end
    
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås    ÇkÇèÇÅÇÑÉCÉxÉìÉg    Åt              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    If Rv_Call3 = "" Then   ' 2006/03/16 Insert
    '   #------------------#
    '   # ìÒ èd ãN ìÆ ñh é~ #
    '   #------------------#
        If Not Dupli_Start Then End
    '   #------------------#
    '   # èâ ä˙ ì‡ óe éÊ ìæ #
    '   #------------------#
        If Not IniGet Then End
    End If                  ' 2006/03/16 Insert
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
    'Dummy2.Left = -1000                                        2009/07/22
    'Dummy3.Left = -1000                                        2009/07/22
    '+----------------------------------- 2004/12/13 Insert
    'DummyKS.Left = -1000                                       2009/07/22
    'DummyKE.Left = -1000                                       2009/07/22
    '+----------------------------------- 2004/12/13 Insert End
    
    Timer1.Interval = 1000   ' éûçèï\é¶ä‘äu
    Timer1.Enabled = True    ' éûçèï\é¶äJén
    Timer2.Interval = 500    ' èàóùãÊï™ì_ñ≈ä‘äu(äÓñ{)
    Timer2.Enabled = False   ' èàóùãÊï™ì_ñ≈ã÷é~(äÓñ{)
    Timer3.Interval = 500    ' èàóùãÊï™ì_ñ≈ä‘äu(éÊà¯êÊ)
    Timer3.Enabled = False   ' èàóùãÊï™ì_ñ≈ã÷é~(éÊà¯êÊ)
    
    mGotColor = B2Op_Hinmoku(1).BackColor
    mLostColor = B2Op_Hinmoku(0).BackColor
    B2Op_Hinmoku(1).BackColor = mLostColor
    B2Op_Hinmoku(1).Value = True
    Call B2Op_Hinmoku_Click(1)
    B2Op_Hinmoku(1).BackColor = mLostColor
    B2Op_Hinmoku(1).Value = False
    ' ñºèÃÉ}ÉXÉ^ämîF
    Call MeisyoInit
    '+--------------------------------+
    '+  ÉOÉäÉbÉhèâä˙ê›íË(éÊà¯êÊèÓïÒ)
    '+--------------------------------+
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 42
        .InitGet = False
        .Init
    End With
    Call GridInit
' =============================================*** 2009/07/22 Start
'    '+--------------------------------+
'    '+  ÉOÉäÉbÉhèâä˙ê›íË(çÏã∆èÓïÒ)
'    '+--------------------------------+
'    With CisVsGridS
'        Set .GridObj = vsGridS
'        .NoSpaceCol = 1
'        .NumberSet = NoSpace
'        .FixedHeight = 350
'        .PatanMax = 1
'        .SelectCol = 5
'        .InitGet = False
'        .Init
'    End With
'    Call GridInitS
'    '+------------------------------------------- insert 2004/12/13
'    '+--------------------------------+
'    '+  ÉOÉäÉbÉhèâä˙ê›íË(ã@äBèÓïÒ)
'    '+--------------------------------+
'    With CisVsGridK
'        Set .GridObj = VsGridK
'        .NoSpaceCol = 1
'        .NumberSet = NoSpace
'        .FixedHeight = 350
'        .PatanMax = 1
'        .SelectCol = 5
'        .InitGet = False
'        .Init
'    End With
'    Call GridInitK
'    '+------------------------------------------- insert 2004/12/13
    '+--------------------------------+
    '+  ÉOÉäÉbÉhèâä˙ê›íË(çHíˆèÓïÒ)
    '+--------------------------------+
    With CisVsGrid2
        Set .GridObj = vsGrid2
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 350
        .PatanMax = 1
        .SelectCol = 9
        .InitGet = False
        .Init
    End With
    Call GridInit2
'---------------------------------------------------' 2009.10.16 del start
'    '+--------------------------------+
'    '+  ÉOÉäÉbÉhèâä˙ê›íË(çHíˆè⁄ç◊èÓïÒ)
'    '+--------------------------------+
'    With CisVsGrid3
'        Set .GridObj = CXM0035.vsGrid3
'        .NoSpaceCol = 1
'        .NumberSet = NoSpace
'        .FixedHeight = 350
'        .PatanMax = 1
'        .SelectCol = 23
'        .InitGet = False
'        .Init
'    End With
'    Call GridInit3
'    '+--------------------------------+
'    '+  ÉOÉäÉbÉhèâä˙ê›íË(æØƒïiî‘èÓïÒ)
'    '+--------------------------------+
'    With CisVsGrid4
'        Set .GridObj = CXM0035.vsGrid4
'        .NoSpaceCol = 1
'        .NumberSet = NoSpace
'        .FixedHeight = 350
'        .PatanMax = 1
'        .SelectCol = 4
'        .InitGet = False
'        .Init
'    End With
'    Call GridInit4
'
'    ' ç›å…ãÊï™ÅiçHíˆÅj
'    Set CisComboKZaiko = New CisCombo
'    With CisComboKZaiko
'        Set .Connect = CisDB
'        Set .CombObj = CXM0035.B3Comb_KZaiko
'        .Control = "ñºèÃÉ}ÉXÉ^"
'        .ComboName = "ç›å…ãÊï™"
'        .Ess = True
'        .IniPath = gIniExe
'        .SetSpace = 50
'        .Ess = True
'        .ComboInit
'    End With
'    ' ç›å…ãÊï™ç≈ëÂÅiçHíˆÅj
'    Set CisComboKMZaiko = New CisCombo
'    With CisComboKMZaiko
'        Set .Connect = CisDB
'        Set .CombObj = CXM0035.B3Comb_KMZaiko
'        .Control = "ñºèÃÉ}ÉXÉ^"
'        .ComboName = "ç›å…ãÊï™"
'        .Ess = True
'        .IniPath = gIniExe
'        .SetSpace = 50
'        .Ess = True
'        .ComboInit
'    End With
''''
''''    '
''''    Back_Grid4.Visible = False
'---------------------------------------------------' 2009.10.16 del end
' =============================================*** 2009/07/22 End
    '+---------------------+
    '+ ∫›ŒﬁŒﬁØ∏Ω èâä˙ê›íË  +
    '+---------------------+
    ' éËîzãÊï™
    Set CisComboTehai = New CisCombo
    With CisComboTehai
        Set .Connect = CisDB
        Set .CombObj = B3Comb_Tehai
        .Control = "ñºèÃÉ}ÉXÉ^"
        .ComboName = "éËîzãÊï™"
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .Ess = True
        .ComboInit
    End With
    ' ç›å…ãÊï™
    Set CisComboZaiko = New CisCombo
    With CisComboZaiko
        Set .Connect = CisDB
        Set .CombObj = B1Comb_Zaiko
        .Control = "ñºèÃÉ}ÉXÉ^"
        .ComboName = "ç›å…ãÊï™"
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .Ess = True
        .ComboInit
    End With

'''2007/12/11DEL    ' î[ïièëå`ë‘(î≠íç)
'''2007/12/11DEL    Set CisComboHNou = New CisCombo
'''2007/12/11DEL    With CisComboHNou
'''2007/12/11DEL        Set .Connect = CisDB
'''2007/12/11DEL        Set .CombObj = B3Comb_Nouhin
'''2007/12/11DEL        .Control = "ñºèÃÉ}ÉXÉ^"
'''2007/12/11DEL        .ComboName = "î[ïièëå`ë‘î≠íç"
'''2007/12/11DEL        .Ess = False
'''2007/12/11DEL        .IniPath = gIniExe
'''2007/12/11DEL        .SetSpace = 50
'''2007/12/11DEL        .ComboInit
'''2007/12/11DEL    End With
'''2007/12/11DEL    ' éÛóÃèëå`ë‘
'''2007/12/11DEL    Set CisComboHJyu = New CisCombo
'''2007/12/11DEL    With CisComboHJyu
'''2007/12/11DEL        Set .Connect = CisDB
'''2007/12/11DEL        Set .CombObj = B3Comb_Jyuryo
'''2007/12/11DEL        .Control = "ñºèÃÉ}ÉXÉ^"
'''2007/12/11DEL        .ComboName = "éÛóÃèëå`ë‘"
'''2007/12/11DEL        .Ess = False
'''2007/12/11DEL        .IniPath = gIniExe
'''2007/12/11DEL        .SetSpace = 50
'''2007/12/11DEL        .ComboInit
'''2007/12/11DEL    End With
'''2007/12/11DEL    ' î[ïièëå`ë‘(éÛíç)
'''2007/12/11DEL    Set CisComboJNou = New CisCombo
'''2007/12/11DEL    With CisComboJNou
'''2007/12/11DEL        Set .Connect = CisDB
'''2007/12/11DEL        Set .CombObj = B3Comb_NouhinJ
'''2007/12/11DEL        .Control = "ñºèÃÉ}ÉXÉ^"
'''2007/12/11DEL        .ComboName = "î[ïièëå`ë‘î≠íç"
'''2007/12/11DEL        .Ess = False
'''2007/12/11DEL        .IniPath = gIniExe
'''2007/12/11DEL        .SetSpace = 50
'''2007/12/11DEL        .ComboInit
'''2007/12/11DEL    End With
'''2007/12/11DEL    ' éÛóÃèëå`ë‘(éÛíç)
'''2007/12/11DEL    Set CisComboJJyu = New CisCombo
'''2007/12/11DEL    With CisComboJJyu
'''2007/12/11DEL        Set .Connect = CisDB
'''2007/12/11DEL        Set .CombObj = B3Comb_JyuryoJ
'''2007/12/11DEL        .Control = "ñºèÃÉ}ÉXÉ^"
'''2007/12/11DEL        .ComboName = "éÛóÃèëå`ë‘éÛíç"
'''2007/12/11DEL        .Ess = False
'''2007/12/11DEL        .IniPath = gIniExe
'''2007/12/11DEL        .SetSpace = 50
'''2007/12/11DEL        .ComboInit
'''2007/12/11DEL    End With
    ' ñæç◊èëî≠çs
    Set CisComboHMei = New CisCombo
    With CisComboHMei
        Set .Connect = CisDB
        Set .CombObj = B3Comb_Meaisai
        .Control = "ñºèÃÉ}ÉXÉ^"
        .ComboName = "ñæç◊èëî≠çs"
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .ComboInit
    End With
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
    Set CisComboJsya = New CisCombo
    With CisComboJsya
        Set .Connect = CisDB
        Set .CombObj = B3Comb_Jisya
        .Control = "ñºèÃÉ}ÉXÉ^"
        .ComboName = "é©é–çHãÊ"
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .ComboInit
    End With
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
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
'-----------------------------------------------------------' 2009.10.20 add start
    PosiSave0X = Back_Hachu(0).Left     ' Ç©åWêîÅAî‰ó¶ÇÃà íuÇãLâØ
    PosiSave0Y = Back_Hachu(0).Top
    PosiSave1X = Back_Hachu(1).Left
    PosiSave1Y = Back_Hachu(1).Top
    PosiSave1W = Back_Hachu(1).Width
'-----------------------------------------------------------' 2009.10.20 add end
    '================================*** 2009/07/22 Start
''''    Back_Kotei.Visible = False                          ' 2009.10.16 del
    PB_Kotei.Visible = False
    '================================*** 2009/07/22 End

    If Rv_Call3 = "" Then
        Me.Show
    End If
End Sub
'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    ÅsÇ`ÇÉÇîÇâÇñÇÅÇîÇÖ ÉCÉxÉìÉgÅt                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

    If Rv_Call3 <> "" Then
        Call Disp_Call
        PB_END.BtnCaption = "ñﬂÇÈ"
    End If

    ' ≤“∞ºﬁäiî[êÊéÊìæ
    With CisFun
'-----------------------------------------------< 2006/02/02 ADD START >
        .INI_FileID = gIniExe & gwsIniName
'-----------------------------------------------< 2006/02/02 ADD START >
        .INI_Section = "Option"
        .INI_Key = "Image":       .INI_Default = gIniExe
        .GetIni
        mImage = RTrim(.INI_String)
'-----------------------------------------------< 2006/02/02 ADD START >
        .INI_FileID = gIniExe & gIniName
'-----------------------------------------------< 2006/02/02 ADD START >
    End With
    If Right(mImage, 1) <> "\" Then mImage = mImage & "\"
'
    FormAct = Not FormAct
' ïiî‘ï“èWèâä˙ílÅ@éÊìæ
    Call EditHinbnInit

'-----     -----     -----     -----     -----     < 2006/02/16 ADD START >
'É|ÉPÉbÉgç≈ëÂêîê›íËílÅ@älìæ
    maxPNo = 60
    Call MeisyoGet("É|ÉPÉbÉgç≈ëÂêî", "1")
    If SYM.êîéö1 <> 0 Then maxPNo = SYM.êîéö1
'-----     -----     -----     -----     -----     < 2006/02/16 ADD  END  >

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
                    Call PB_ADD_Click    'ÅyêVãKÅz
                End If
           '+----------------------------------------------------- 2004/12/13 Delete
'                If ProcHB = "B3" And iB3_No.Visible Then
'                   On Error Resume Next
'                   If ActiveControl.Name = "iB3_No" Then
'                      iB3_Seban.SetFocus
'                   Else
'                      iB3_No.SetFocus
'                   End If
'                   On Error GoTo 0
'                End If
           '+----------------------------------------------------- 2004/12/13 Delete End
'================================================ *** 2009/07/22 Start
'           '+----------------------------------------------------- 2004/12/13 Insert
'                If ProcHB = "B3" And iB3_Kikai.Visible Then
'                   On Error Resume Next
'                   If ActiveControl.Name = "iB3_Kikai" Then
'                      iB3_Seban.SetFocus
'                   Else
'                      iB3_Kikai.SetFocus
'                   End If
'                   On Error GoTo 0
'                End If
''''                If PB_KADD.Visible Then
''''                    Call PB_KADD_Click    'ÅyêVãKÅz
''''                End If
'================================================ *** 2009/07/22 End
           Case vbKeyF2:
                If PB_MNT.Visible Then
                    Call PB_MNT_Click    'ÅyèCê≥Åz
                End If
                '=================================== *** 2009/07/22 Start
''''                If PB_KMNT.Visible Then
''''                    Call PB_KMNT_Click    'ÅyèCê≥Åz
''''                End If
                '=================================== *** 2009/07/22 End
'================================================= *** 2009/07/22 Start
'                If ProcHB = "B3" And iB3_No.Visible Then
'                   On Error Resume Next
'                   If ActiveControl.Name = "iB3_No" Then
'                      iB3_Seban.SetFocus
'                   Else
'                      iB3_No.SetFocus
'                   End If
'                   On Error GoTo 0
'                End If
'                If ProcHB = "B3" Then                                  '2009/07/22
''''                If ProcHB = "B3" And Back_Kotei.Visible = False Then    '2009/07/22
                   On Error Resume Next
                   iB3_Seban.SetFocus
                   On Error GoTo 0
''''                End If
'================================================= *** 2009/07/22 End

           '+----------------------------------------------------- 2004/12/13 Insert End
           '+----------------------------------------------------- 2004/12/13 Delete
           'Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    'ÅyèCê≥Åz
           '+----------------------------------------------------- 2004/12/13 Delete End
'================================================================= 2009/07/22 Start
'           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    'ÅyçÌèúÅz
           Case vbKeyF3
                                If PB_DEL.Visible Then Call PB_DEL_Click    'ÅyçÌèúÅz
''''                                If PB_KDEL.Visible Then Call PB_KDEL_Click    'ÅyçÌèúÅz
'================================================================= 2009/07/22 End
           Case vbKeyF10:       If PB_Zairyo.Visible Then Call PB_Zairyo_Click  'Åyçﬁóøä÷òAïtÇØÅz
'           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  'ÅyåüçıÅz
           Case vbKeyF7
                               If PB_Look.Visible Then Call PB_Look_Click  'ÅyåüçıÅz
''''                               If PB_KLook.Visible Then Call PB_KLook_Click 'ÅyåüçıÅz
           Case vbKeyF8:       If PB_ImageDisp.Visible Then Call PB_ImageDisp_Click  'Åy≤“∞ºﬁÅz
           Case vbKeyF9:       If PB_Chg.Visible Then Call PB_CHG_Click    'ÅyåüçıÅz
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  'Åyå„ëﬁÅz
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  'ÅyëOêiÅz
'================================================================= 2009/07/22 Start
           'Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    'ÅyéÊè¡Åz
           'Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    'ÅyèIóπÅz
           Case vbKeyEscape
                               If PB_CAN.Visible Then Call PB_CAN_Click    'ÅyéÊè¡Åz
''''                               If PB_KCAN.Visible Then Call PB_KCAN_Click    'ÅyéÊè¡Åz
           Case vbKeyF12
                               If PB_END.Visible Then Call PB_END_Click    'ÅyèIóπÅz
''''                               If PB_KEND.Visible Then Call PB_KEND_Click    'ÅyèIóπÅz
'================================================================= 2009/07/22 End
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      'Åyì¸óÕÅz
                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
                                  Dummy.Enabled = False
                               End If
                               '============================================ *** 2009/07/22 Start
''''                               If PB_KENT.Visible Then                      'Åyì¸óÕÅz
''''                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
''''                                  Call PB_KENT_Click
''''                                  Dummy.Enabled = False
''''                               End If
                               '============================================ *** 2009/07/22 End
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
           Case vbKeyF4:       If PB_UTN.Visible Then Call PB_UTN_Click    'ÅyîÑè„íPâøÅz
           Case vbKeyF5:       If PB_KTN.Visible Then Call PB_KTN_Click    'Åyçwì¸íPâøÅz
           Case vbKeyF6:       If PB_STN.Visible Then Call PB_STN_Click    'ÅyéxããíPâøÅz
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
           
           Case vbKeyF11:      If PB_Kotei.Visible Then Call PB_Kotei_Click    'ÅyçHèáÅz        '2009/07/22
           
           Case Else:          Exit Function
    End Select
    Key_Acc = True
End Function
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås QueryUnloadÉCÉxÉìÉg Åt                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
' *CHG*  2004/04/05
'    Cancel = 1
    If Rv_Call3 = "" Then
       Cancel = 1
    End If
    
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      åƒèoÇ≥ÇÍÇΩèÍçáÇÃèâä˙ï\é¶
'+----------------------------------------------------------+
Private Sub Disp_Call()
    SyoriKB = "MNT"
    Call SyoriName(SyoriKB)
    
    iH1_Hinbn = RV_Hinbn
    Call ReturnPress

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/03/31 DELETE START
'    PB_Chg.Tag = "TOR"
'    Call PB_Chg_Click
'    PB_Chg.Visible = False
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/03/31 DELETE END
End Sub

'-----------------------------------------------------------' 2009.10.20 add start
'+-----------------------------+                            '
'+  ë„ï\ìæà”êÊ      iB1_DaiCD
'+-----------------------------+
Private Sub iB1_DaiCD_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_DaiCD_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B1lb_DaiNM = ""
    If Trim(iB1_DaiCD.Text) <> "" Then
       If TorNmGet(iB1_DaiCD, 3) Then
          B1lb_DaiNM = TRM_RName
''''       Else
''''          If ProcHB = "B1" Then
''''             With CisFun
''''                .MB_Lines = 4
''''                .MB_MSG(2) = "    ë„ï\ìæà”êÊ  ñ¢ìoò^              "
''''                .MB_Title = ""
''''                .MB_Button = Error
''''                .MBOX
''''             End With
''''             iB1_DaiCD.SetFocus
''''          End If
       End If
    End If
End Sub
'+-----------------------------+                            '
'+  É|ÉCÉìÉg        iB3_Point
'+-----------------------------+
Private Sub iB3_Point_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_Point_LostFocus()
    
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    Image1.Picture = Nothing
    If iB3_Point <> "" Then
       POM.É|ÉCÉìÉg = RTrim(iB3_Point.Text)
       If POMRead("") = True Then
          On Error Resume Next
          Set Image1.Picture = LoadPicture(gIniExe & "Image\" & RTrim(POM.ÉCÉÅÅ[ÉW))
''''       Else
''''          If ProcHB = "B3" Then
''''             With CisFun
''''               .MB_Lines = 4
''''               .MB_MSG(2) = "    É|ÉCÉìÉg  ñ¢ìoò^                "
''''               .MB_Title = ""
''''               .MB_Button = Error
''''               .MBOX
''''             End With
''''             iB3_Point.SetFocus
''''          End If
       End If
    End If
End Sub
'+-----------------------------+                            '
'+  ï‘ãpèÍèäÅ@      iB1_HenCD
'+-----------------------------+
Private Sub iB3_HenCD_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_HenCD_LostFocus()
    Dim wSel    As String

    B3lb_HenNM = ""
    If Trim(iB3_HenCD.Text) <> "" Then
       wSel = "SELECT * FROM ñºèÃÉ}ÉXÉ^"
       wSel = wSel & " WHERE ãÊï™ñºèÃ = 'ï‘ãpèÍèä'"
       wSel = wSel & "   AND íl       = '" & Trim(iB3_HenCD.Text) & "'"

       If SYMRead(wSel) = True Then
          B3lb_HenNM = SYM.ílñºèÃ
''''       Else
''''          If ProcHB = "B3" Then
''''             With CisFun
''''               .MB_Lines = 4
''''               .MB_MSG(2) = "    ï‘ãpèÍèä  ñ¢ìoò^                "
''''               .MB_Title = ""
''''               .MB_Button = Error
''''               .MBOX
''''             End With
''''             iB3_HenCD.SetFocus
''''          End If
       End If
    End If
End Sub
'+-----------------------------+                            '
'+  å„çHíˆÅ@Å@      iB1_AtoCD
'+-----------------------------+
Private Sub iB3_AtoCD_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_AtoCD_LostFocus()
    Dim wSel    As String

    If CisFun.Val2(iB3_AtoCD.Text) = 0 Then
       iB3_AtoCD.Text = ""
    Else
       iB3_AtoCD.Text = Format(CisFun.Val2(iB3_AtoCD.Text), "000")
    End If

    B3lb_AtoNM = ""
    If Trim(iB3_AtoCD.Text) <> "" Then
       wSel = "SELECT TM.ó™èÃ, M2.ílñºèÃ ê≥éÆñºèÃ"                                     ' , tm.*, m1.*
       wSel = wSel & " FROM  éÊà¯êÊÉ}ÉXÉ^ TM"
       wSel = wSel & "       LEFT OUTER JOIN ñºèÃÉ}ÉXÉ^ M1"
       wSel = wSel & "              ON  M1.ãÊï™ñºèÃ = 'çHíˆÇ©ÇÒÇŒÇÒî≠çsÉpÉ^Å[Éì'"
       wSel = wSel & "              AND M1.íl       = '" & Trim(iB3_AtoCD.Text) & "'"  ' å„çHíˆ
       wSel = wSel & "              AND M1.ílñºèÃ   = '" & Trim(iB2_Torcd.Text) & "'"  ' éÊà¯êÊ
       wSel = wSel & "       LEFT OUTER JOIN ñºèÃÉ}ÉXÉ^ M2"
       wSel = wSel & "              ON  M2.ãÊï™ñºèÃ = 'ï‘ãpèÍèä'"
       wSel = wSel & "              AND M2.íl       = M1.âpêîéö3"                      ' ï‘ãpèÍèä1
       wSel = wSel & " WHERE TM.éÊà¯êÊCD   = M1.âpêîéö2"                               ' ïîèê1
       wSel = wSel & "   AND TM.éÊà¯êÊãÊï™ = 1"                                        ' éÊà¯êÊãÊï™ = î≠íç  ' 2009.10.23 add
'''''''wSel = wSel & "   AND TM.é–ì‡ãÊï™   = 1"                                        ' é–ì‡ãÊï™   = é–ì‡  ' 2009.10.23 del

       If TRMRead(wSel) = True Then
          B3lb_AtoNM = Trim(TRM.ó™èÃ) & " " & Trim(TRM.ê≥éÆñºèÃ)   ' ê≥éÆñºèÃÇ…ï‘ãpèÍèä1ÇÉZÉbÉgÇµÇƒÇ¢ÇÈ
''''       Else
''''          If ProcHB = "B3" Then
''''             With CisFun
''''               .MB_Lines = 4
''''               .MB_MSG(2) = "    å„çHíˆ  ñ¢ìoò^                  "
''''               .MB_Title = ""
''''               .MB_Button = Error
''''               .MBOX
''''             End With
''''             iB3_AtoCD.SetFocus
''''          End If
       End If
    End If
End Sub
'-----------------------------------------------------------' 2009.10.20 add end

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
'        Call CisFrm.UnLoadDisp
        Unload CXM0030
        DoEvents
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
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call KJ_Delete_Chk(0)                     '2009/07/22
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKB)
      iH1_Hinbn.SetFocus
      GoTo PB_CAN_Ed
   End If
   If ProcHB = "B2" Then
      Call KJ_Delete_Chk(1)                     '2009/07/22
      
      SyoriKB2 = ""
      Call SyoriName2(SyoriKB2)
      ProcHB = "B1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B2")
      Call HeadBodyClear("B3")
      If Trim(vsGrid1.TextMatrix(1, 1)) <> "" Then
         vsGrid1.SetFocus
      Else
         iB1_Syakei.SetFocus
      End If
      GoTo PB_CAN_Ed
   End If
   If ProcHB = "B3" Then
      Call KJ_Delete_Chk(1)                     '2009/07/22
      
      RowColFlg = True
      Call CisVsGrid.vsColor(è¡ãéëSïî)
      ProcHB = GridBack
      RowColFlg = False
      If ProcHB = "B1" Then
         Call HeadBodyClear("B2")
      End If
      Call HeadBodyClear("B3")
      Call DispChange(ProcHB)
      If ProcHB = "B1" Then
         SyoriKB2 = ""
         Call SyoriName2(SyoriKB2)
         Call DispChange(ProcHB)
         vsGrid1.SetFocus
      Else
         Call SyoriName2(SyoriKB2)
         iB2_Torcd.SetFocus
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
    If ProcHB = "H1" Then
        If SyoriKB <> "ADD" Then
            SyoriKB = "ADD"
            Call SyoriName(SyoriKB)
        End If
        iH1_Hinbn.SetFocus
    Else
        If SyoriKB2 <> "ADD" Then
            SyoriKB2 = "ADD"
            Call SyoriName2(SyoriKB2)
        End If
        iB2_Torcd.SetFocus
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
        iH1_Hinbn.SetFocus
    Else
        If SyoriKB2 <> "MNT" Then
            SyoriKB2 = "MNT"
            Call SyoriName2(SyoriKB2)
        End If
        iB2_Torcd.SetFocus
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
        iH1_Hinbn.SetFocus
    Else
        If SyoriKB2 <> "DEL" Then
            SyoriKB2 = "DEL"
            Call SyoriName2(SyoriKB2)
        End If
        iB2_Torcd.SetFocus
    End If
End Sub
'+----------------------------+
'+    å„  ëﬁ  ÉL Å[(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = " SELECT *"
    gSL_Select = gSL_Select & " FROM ïiî‘É}ÉXÉ^ "
    gSL_Select = gSL_Select & " WHERE ïiî‘ < '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " ORDER BY ïiî‘ Desc "
    Call DBInput("Back")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Hinbn.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    ëO  êi  ÉL Å[(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = " SELECT *"
    gSL_Select = gSL_Select & " FROM ïiî‘É}ÉXÉ^ "
    gSL_Select = gSL_Select & " WHERE ïiî‘ > '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " ORDER BY ïiî‘ "
    Call DBInput("Next")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Hinbn.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    êÿÅ@ë÷  ÉL Å[(F09)       +
'+----------------------------+
Private Sub PB_CHG_Click()
    If PB_Chg.Tag = "TOR" Then
       If Not Body1Chk Then Exit Sub
       Call HeadBodyClear("B2")
       Call HeadBodyClear("B3")
       ProcHB = "B2"
       Call DispChange(ProcHB)
       SyoriKB2 = "ADD"
       Call SyoriName2(SyoriKB2)
       B2Op_Hinmoku(1).Value = True
       iB2_Torcd.SetFocus
    Else
       Call PB_CAN_Click
    End If
End Sub
'+----------------------------+
'+    ≤“∞ºﬁÃß≤ŸéwíË           +
'+----------------------------+
Private Sub PB_Image_Click()
     On Error Resume Next
     CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly
'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2006/12/19 Update Satart
     'CommonDialog1.FileName = iB1_Image
     CommonDialog1.FileName = mImage & iB1_Image
'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2006/12/19 Update End
     gStr = "Ç∑Ç◊ÇƒÇÃÀﬂ∏¡¨Ãß≤Ÿ| *.bmp;;*dib;*.gif;*.jpg;*.wmf;*.emf"
     gStr = gStr & "|ÀﬁØƒœØÃﬂ    (*.bmp;*.dib)|*.bmp;*.dib"
     gStr = gStr & "|GIF ≤“∞ºﬁ   (*.gif)|*.gif"
     gStr = gStr & "|JPEG ≤“∞ºﬁ  (*.jpg)|*.jpg"
     gStr = gStr & "|“¿Ãß≤Ÿ      (*.wmf;*.emf)|*.wmf;*.emf"
     gStr = gStr & "|Ç∑Ç◊ÇƒÇÃÃß≤Ÿ (*.*)|*.*"
     CommonDialog1.Filter = gStr
    ' à»â∫ÇÃø∞Ω∫∞ƒﬁÇÕÅA<∑¨›æŸ>Œﬁ¿›Ç™∏ÿØ∏Ç≥ÇÍÇΩÇ∆Ç´Ç…é¿çsÇ≥ÇÍÇ‹Ç∑ÅB
     CommonDialog1.CancelError = True
     CommonDialog1.ShowOpen
     If Err = 32755 Then
    '<∑¨›æŸ>Œﬁ¿›Ç™âüÇ≥ÇÍÇΩÇ∆Ç´Ç…¥◊∞∫∞ƒﬁÇï‘Ç∑ÇÊÇ§Ç…Ç∑ÇÈê›íËÇ≈Ç∑ÅB
    'ïKóvÇ»ø∞Ω∫∞ƒﬁÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB
        Exit Sub
     End If
     If Err = 20477 Then
        With CisFun
           .MB_Lines = 4
           .MB_MSG(2) = "    ÉtÉ@ÉCÉãñºÇ™ïsê≥Ç≈Ç∑ÅB            "
           .MB_Title = ""
           .MB_Button = Error
           .MBOX
        End With
        Exit Sub
     End If
'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2006/12/19 Update Satart
'     If Trim(CommonDialog1.FileName) <> "" Then
'        With CisFun
'            .FE_FullPath = CommonDialog1.FileName
'            .GetPathName
'            iB1_Image = .FE_File
'        End With
'        Call iB1_Image_LostFocus
'     End If
    Static sw_UNCImgF           As String
    Static sw_UNCImgS           As String
    Static sw_Img               As String

    If Left$(mImage, 2) = "\\" Then
        sw_UNCImgF = mImage
    Else
'       * UNC *
        With CisFun
            .FE_FullPath = mImage
            .GetUNCPath
        
            sw_UNCImgF = RTrim$(.FE_UNCPath)
        End With
    End If
     
    If Trim(CommonDialog1.FileName) <> "" Then

        If Left$(Trim(CommonDialog1.FileName), 2) = "\\" Then
            sw_UNCImgS = Trim(CommonDialog1.FileName)
        Else
    '       * UNC *
            With CisFun
                .FE_FullPath = Trim(CommonDialog1.FileName)
                .GetPathName
                sw_Img = .FE_File
                
                .GetUNCPath
                sw_UNCImgS = RTrim$(.FE_UNCPath)
            End With
        End If

        sw_Img = Replace(UCase(sw_UNCImgS), UCase(sw_UNCImgF), "")
        
        If InStr(sw_Img, ":") <> 0 Or _
           InStr(sw_Img, "\\") <> 0 Then
            With CisFun
                .FE_FullPath = CommonDialog1.FileName
                .GetPathName
                sw_Img = .FE_File
            End With
        End If
        
        iB1_Image = sw_Img
        Call iB1_Image_LostFocus
    End If
'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2006/12/19 Update End
     iB1_Image.SetFocus
     On Error GoTo 0
End Sub
'+----------------------------+
'+    ≤“∞ºﬁÃß≤ŸéwíË           +
'+----------------------------+
Private Sub PB_ImageDisp_Click()
    If Trim(iB1_Image) = "" Then
        With CisFun
           .MB_Lines = 4
           .MB_MSG(2) = "    ÉtÉ@ÉCÉãñºÇ™ê›íËÇ≥ÇÍÇƒÇ¢Ç‹ÇπÇÒ            "
           .MB_Title = ""
           .MB_Button = Error
           .MBOX
        End With
        Exit Sub
    End If
    With CisFun
        .ShellApl = mImage & iB1_Image
        .ShellAplPara = ""
        .ShellTaskEX NoFreeze
    End With
End Sub
'+----------------------------+
'+    çﬁóøämîF                 +
'+----------------------------+
Private Sub PB_Zairyo_Click()
    If Not Body1Chk Then Exit Sub
    If Not DBPut(False) Then Exit Sub

    RV_Left = 0
    RV_Top = 0
    Rv_Call2 = "HIN"
    RV_Hinbn = RTrim(iH1_Hinbn)

    CXM0110.Show vbModal
    Unload CXM0110
    Set CXM0110 = Nothing

    Call ZairyoKensu
End Sub
'+--------------------------------+
'+      çﬁóøåèêîï\é¶
'+--------------------------------+
Private Sub ZairyoKensu()
    ' çﬁóøèÓïÒ
    gSL_Select = "Select Count(*) åèêî from çﬁóøïiî‘É}ÉXÉ^ "
    gSL_Select = gSL_Select & " where égópïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    If IRNRead(gSL_Select, 1) Then
       B1lb_Zairyo = Format(IRN.åèêî, "##")
    Else
       B1lb_Zairyo = ""
    End If
    
End Sub
'+----------------------------+
'+    åü  çı  ÉL Å[(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iB1_KUtsuwa" Then Call Look_Kubun(4): Exit Sub
    If PB_Look.Tag = "iB1_KataTor" Then Call Look_Torcd: Exit Sub
    '''If PB_Look.Tag = "iB1_Kotei" Then Call Look_Kubun(0): Exit Sub
    If PB_Look.Tag = "iB1_Hinku" Then Call Look_Kubun(9): Exit Sub
    '''If PB_Look.Tag = "iB1_NKotei" Then Call Look_Kubun(1): Exit Sub

    If PB_Look.Tag = "iB1_Tanto" Then Call Look_Kubun(2): Exit Sub
    If PB_Look.Tag = "iB2_Torcd" Then Call Look_SSTorcd: Exit Sub
    If PB_Look.Tag = "iB2_Ukeir" Then Call Look_UTorcd(0): Exit Sub
    If PB_Look.Tag = "iB3_Youki1" Then Call Look_Kubun(5): Exit Sub
    If PB_Look.Tag = "iB3_Youki2" Then Call Look_Kubun(6): Exit Sub
    If PB_Look.Tag = "iB3_KTorcd" Then Call Look_KNTorcd(0): Exit Sub
    If PB_Look.Tag = "iB3_NTorcd" Then Call Look_KNTorcd(1): Exit Sub
    If PB_Look.Tag = "iB3_DaikoK" Then Call Look_KNTorcd(2): Exit Sub
    If PB_Look.Tag = "iB3_Ukeir" Then Call Look_UTorcd(1): Exit Sub
    If PB_Look.Tag = "iB3_DenSyu" Then Call Look_Kubun(7): Exit Sub
    If PB_Look.Tag = "iB3_DenSyuJ" Then Call Look_Kubun(8): Exit Sub
    If PB_Look.Tag = "iB3_Sagyo" Then Call Look_Kubun(3): Exit Sub
    If PB_Look.Tag = "iB3_SyukaK" Then Call Look_Syanai: Exit Sub
    If PB_Look.Tag = "iB1_DaiCD" Then Call Look_KNTorcd(3): Exit Sub        ' 2009.10.20 add
    If PB_Look.Tag = "iB3_HenCD" Then Call Look_Kubun(13):  Exit Sub        ' 2009.10.21 add
    If PB_Look.Tag = "iB3_AtoCD" Then Call Look_AtoCD: Exit Sub             ' 2009.10.21 add
    If PB_Look.Tag = "iB3_Point" Then Call Look_Point: Exit Sub
    '+------------------------------------------------------------- 2004/12/13 Insert
    'If PB_Look.Tag = "iB3_Kikai" Then Call Look_Kikai: Exit Sub                            2009/07/22
    '+------------------------------------------------------------- 2004/12/13 Insert End
    '+------------------------------------------------------------- 2005/09/30 Insert
    If PB_Look.Tag = "iB3_PNo" Then Call Look_PNO: Exit Sub
    '+------------------------------------------------------------- 2004/09/30 Insert End
    If PB_Look.Tag = "iB3_HYmd" Then
       iB3_HYmd.ShowCalender
       Exit Sub
    End If
    If PB_Look.Tag = "iB3_YSYmd" Then
       iB3_YSYmd.ShowCalender
       Exit Sub
    End If
    If PB_Look.Tag = "iB3_YEYmd" Then
       iB3_YEYmd.ShowCalender
       Exit Sub
    End If
End Sub
'-------------------------------------------' 2009.10.21 add start
'+----------------------------+
'+       å„çHíˆåüçı           +
'+----------------------------+
Private Sub Look_AtoCD()

    RV_Left = 0
    RV_Top = 0
    
    RV_Torcd = Trim(iB2_Torcd)  ' ïîèê  (éÊà¯êÊ)
    RV_TorcdK = Trim(iB3_HenCD) ' ï‘ãpèÍèä
    RV_Name = Trim(B2lb_Torcd)  ' ïîèêñº(éÊà¯êÊñº)
    RV_NameR = Trim(B3lb_HenNM) ' ï‘ãpèÍèäñº
    
    CKK0012.Show vbModal
    Unload CKK0012
    Set CKK0012 = Nothing

    If RV_Rtn Then
       iB3_AtoCD = RV_Name      ' å„çHíˆ
      'B3lb_AtoNM = RV_NameR    ' å„çHíˆñº
       Call iB3_AtoCD_LostFocus
    End If
End Sub
'+----------------------------+
'+       É|ÉCÉìÉgåüçı          +
'+----------------------------+
Private Sub Look_Point()

    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"

    CKK3320.Show vbModal
    Unload CKK3320
    Set CKK3320 = Nothing

    If RV_Rtn Then
       iB3_Point = RV_Code
        
        Image1.Picture = Nothing
        If iB3_Point <> "" Then
           POM.É|ÉCÉìÉg = RTrim(iB3_Point.Text)
           If POMRead("") = True Then
              On Error Resume Next
              Set Image1.Picture = LoadPicture(gIniExe & "Image\" & RTrim(POM.ÉCÉÅÅ[ÉW))
           End If
        End If
    
    End If

End Sub
'-------------------------------------------' 2009.10.21 add end

'=============================================== *** 2009/07/22 Start
'+--------------------------------------+
'+    çHèáèÓïÒ√ﬁ∞¿ìoò^/çÌèúämîF         +
'+--------------------------------------+
Private Sub KJ_Delete_Chk(KB As Byte)
    If SyoriKB = "ADD" Or SyoriKB = "MNT" Then
       If KB = 0 Then
           'ïiî‘íPà 
           HNM.ïiî‘ = RTrim(iH1_Hinbn)
           If Not HNMRead("", 1) Then
              With CisDB
                  On Error GoTo KJ_DBDelete_Err1
                  .DBTran (TransBegin) ' ƒ◊›ªﬁ∏ºÆ›äJén
          ' ( çÌèúèàóù )
                  ' ïiî‘çHèáÉ}ÉXÉ^
                  .SQL = "DELETE FROM ïiî‘çHèáÉ}ÉXÉ^ "
                  .SQL = .SQL & " WHERE ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
                  On Error GoTo KJ_DBDelete_Err1
                  .DBExec
                  ' ïiî‘çHèáÉZÉbÉgïiî‘É}ÉXÉ^
                  .SQL = "DELETE FROM ïiî‘çHèáÉZÉbÉgïiî‘É}ÉXÉ^ "
                  .SQL = .SQL & " WHERE ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
                  On Error GoTo KJ_DBDelete_Err1
                  .DBExec
                    
                  .DBTran (TransCommit) ' ƒ◊›ªﬁ∏ºÆ› èIóπ
                   
                  On Error GoTo 0
              End With
              Exit Sub
KJ_DBDelete_Err1:
              CisFun.ErrorBox
              End
           End If
       End If
       
       If KB = 1 Then
           'ïiî‘éÊà¯êÊíPà 
           gSL_Select = "Select * from ïiî‘éÊà¯êÊÉ}ÉXÉ^ "
           gSL_Select = gSL_Select & " Where ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
           gSL_Select = gSL_Select & " And   éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
           gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
           If B2Op_Hinmoku(0) Then
              gSL_Select = gSL_Select & " And   ïiñ⁄ = 0"
           Else
              gSL_Select = gSL_Select & " And   ïiñ⁄ = 1"
           End If
           If Not HTMRead(gSL_Select, 1) Then
              With CisDB
                  On Error GoTo KJ_DBDelete_Err2
                  .DBTran (TransBegin) ' ƒ◊›ªﬁ∏ºÆ›äJén
          ' ( çÌèúèàóù )
                  ' ïiî‘çHèáÉ}ÉXÉ^
                  .SQL = "DELETE FROM ïiî‘çHèáÉ}ÉXÉ^ "
                  .SQL = .SQL & " Where ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
                  .SQL = .SQL & " And   éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
                  .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
                  If B2Op_Hinmoku(0) Then
                     .SQL = .SQL & " And   ïiñ⁄ = 0"
                  Else
                     .SQL = .SQL & " And   ïiñ⁄ = 1"
                  End If
                  On Error GoTo KJ_DBDelete_Err2
                  .DBExec
                  ' ïiî‘çHèáÉZÉbÉgïiî‘É}ÉXÉ^
                  .SQL = "DELETE FROM ïiî‘çHèáÉZÉbÉgïiî‘É}ÉXÉ^ "
                  .SQL = .SQL & " Where ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
                  .SQL = .SQL & " And   éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
                  .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
                  If B2Op_Hinmoku(0) Then
                     .SQL = .SQL & " And   ïiñ⁄ = 0"
                  Else
                     .SQL = .SQL & " And   ïiñ⁄ = 1"
                  End If
                  On Error GoTo KJ_DBDelete_Err2
                  .DBExec
                  
                  .DBTran (TransCommit) ' ƒ◊›ªﬁ∏ºÆ› èIóπ
            
              End With
              Exit Sub
KJ_DBDelete_Err2:
              CisFun.ErrorBox
              End
           End If
       End If
       
    End If
End Sub

'---------------------------------------------------------------------------' 2009.10.16 del start
'''''+--------------------------------------+
'''''+    åü  çı  ÉL Å[ÅFçHíˆèÓïÒ(F7)       +
'''''+--------------------------------------+
''''Private Sub PB_KLook_Click()
''''    If PB_KLook.Tag = "iB3_Kikai" Then Call Look_Kubun(10): Exit Sub
''''    If PB_KLook.Tag = "iB3_KikaiNo" Then Call Look_Kubun(11): Exit Sub
''''    If PB_KLook.Tag = "iB3_Kotei" Then Call Look_Kubun(12): Exit Sub
''''    If PB_KLook.Tag = "iB3_GTorcd" Then Call Look_GTorcd: Exit Sub
''''    If PB_KLook.Tag = "iB3_SetHinbn" Then Call Look_SetHinbn: Exit Sub
''''End Sub
'---------------------------------------------------------------------------' 2009.10.16 del end
'+----------------------------+
'+    çHíˆâÊñ ï\é¶  ÉL Å[     +
'+----------------------------+
Private Sub PB_Kotei_Click()
'-----------------------------------' 2009.10.16 del start
''''    Back_Kotei.Visible = True
''''    PNL_PF.Visible = False
''''    PB_Kotei.Enabled = False
''''    B1_Area1.Enabled = False
''''    B3_Area1.Enabled = False
''''    B3_Area3.Enabled = False
''''    B1_Area2.Enabled = False
''''    SyoriKB_K = "ADD"
''''    SyoriNM_K = "êVãK"
''''    Call B3Comb_KZaiko_Click
''''    Call B3Comb_KMZaiko_Click
''''    iB3_Kojyun.SetFocus
''''    Call CisComboKZaiko.ComboSet
''''    Call CisComboKMZaiko.ComboSet
''''    Call PB_KADD_Click
'-----------------------------------' 2009.10.16 del end
    
'-----------------------------------------------------------------------------------------------' 2009.10.16 add start
   'gSL_Select = SAVE_Select    ' ç≈êVï\é¶Ç≥ÇÍÇƒÇ¢ÇÈçHèáÇï\é¶ÇµÇΩÇ∆Ç´ÇÃSQLÇèâä˙ï\é¶ópÇ…ìnÇ∑
    CXM0035.Show vbModal
    Unload CXM0035
    Set CXM0035 = Nothing
    
''''    If RV_Rtn Then
''''       iB3_NTorcd = RV_TorcdK       '''''''''' Ç‹ÇæâºÇ»ÇÃÇ≈ÅAëóéÛämíËÇµÇΩÇÁèCê≥Ç∑ÇÈÇ±Ç∆
''''       iB3_Ukeir = RV_Ukeir
''''       B3lb_NTorcd = RV_TorRName
''''    End If
    Call Kojyun_SQL_Set
    If HKTRead(gSL_Select, , 1) Then
       Call GridSet2
    End If
'-----------------------------------------------------------------------------------------------' 2009.10.16 add end
End Sub
'+-----------------------------+
'+  iB3_SyanaiK
'+-----------------------------+
Private Sub iB3_SyukaK_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_SyukaK_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B3lb_SyukaKNM = ""
    If TorNmGet(iB3_SyukaK, 6) Then
        B3lb_SyukaKNM = TRM_RName
    End If
End Sub

'+----------------------------+
'+       ïiî‘åüçı            +
'+----------------------------+
Private Sub Look_Hinbn()

    RV_Left = 0
    RV_Top = 0
    RVI_Hinmoku = 9
    RVI_SK = 9
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
       iH1_Hinbn = RV_Hinbn
    End If
End Sub
'+----------------------------+
'+       éÊà¯êÊåüçı            +
'+----------------------------+
Private Sub Look_Torcd()

    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 1
    
    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

    If RV_Rtn Then
       iB1_KataTor = RV_TorcdK
       B1lb_KataTor = RV_TorRName
    End If
End Sub
'+-----------------------------+
'+       éÊà¯êÊåüçı             +
'+-----------------------------+
Private Sub Look_SSTorcd()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    If B2Op_Hinmoku(0).Value Then
       RV_TorKb = 3
    Else
       RV_TorKb = 9
    End If
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iB2_Torcd = RV_TorcdK
        B2lb_Torcd = RV_TorRName
    End If
End Sub
'+----------------------------+
'+       éÊà¯êÊ(é–ì‡)åüçı      +
'+----------------------------+
Private Sub Look_Syanai()

    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 1
    RV_TorKb2 = 1
    RV_Call = "1"
    
    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

    If RV_Rtn Then
       iB3_SyukaK = RV_TorcdK
       B3lb_SyukaKNM = RV_TorRName
    End If
End Sub

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
'+--------------------------------+
'+     îÑè„íPâøÉ{É^ÉìÉNÉäÉbÉN     +
'+--------------------------------+
Private Sub PB_UTN_Click()
    Rv_Call3 = "X"
    RV_HinZai = 0                   ' ïiî‘,çﬁóøãÊï™
    RV_Hinbn = RTrim(iH1_Hinbn)
    RV_Torcd = RTrim(iB2_Torcd)
    RV_SYmd = ""
    RV_EYmd = ""
    CXM0160.Show vbModal
    Unload CXM0160
    Set CXM0160 = Nothing

    If RV_Rtn Then
    End If
    Rv_Call3 = ""
End Sub
'+--------------------------------+
'+     çwì¸íPâøÉ{É^ÉìÉNÉäÉbÉN     +
'+--------------------------------+
Private Sub PB_KTN_Click()
    Rv_Call3 = "X"
    RV_HinZai = 0                   ' ïiî‘,çﬁóøãÊï™
    RV_Hinbn = RTrim(iH1_Hinbn)
    
    Call CisComboTehai.CodeGet
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/12 UPDATE START
'    If B2Op_Hinmoku(1).Value And CisComboTehai.Code = "1" Then
'       RV_Torcd = RTrim(iB3_NTorcd)
'    Else
'       RV_Torcd = RTrim(iB2_Torcd)
'    End If
    RV_Torcd = RTrim(iB2_Torcd)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/12 UPDATE END
    RV_KTorcd = RTrim(iB3_KTorcd)
    RV_SYmd = ""
    RV_EYmd = ""
    Rv_Call3Rtn = 0
    CXM0180.Show vbModal
    Unload CXM0180
    Set CXM0180 = Nothing

    If RV_Rtn Then
    End If
    Rv_Call3 = ""
End Sub
'+--------------------------------+
'+     éxããíPâøÉ{É^ÉìÉNÉäÉbÉN     +
'+--------------------------------+
Private Sub PB_STN_Click()
    Rv_Call3 = "X"
    RV_HinZai = 0                   ' ïiî‘,çﬁóøãÊï™
    RV_Hinbn = RTrim(iH1_Hinbn)
    
    Call CisComboTehai.CodeGet
    If B2Op_Hinmoku(1).Value And CisComboTehai.Code = "1" Then
       RV_Torcd = RTrim(iB3_NTorcd)
    Else
       RV_Torcd = RTrim(iB2_Torcd)
    End If
    RV_SYmd = ""
    RV_EYmd = ""
    CXM0190.Show vbModal
    Unload CXM0190
    Set CXM0190 = Nothing

    If RV_Rtn Then
    End If
    Rv_Call3 = ""
End Sub
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END

'+-----------------------------+
'+       éÊà¯êÊéÛì¸åüçı         +
'+-----------------------------+
Private Sub Look_UTorcd(KB As Integer)
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    If KB = 0 Then
        If B2Op_Hinmoku(0).Value Then
           RV_TorKb = 3
        Else
           RV_TorKb = 9
        End If
        RV_Torcd = iB2_Torcd
    Else
        RV_Torcd = iB3_NTorcd
        If B2Op_Hinmoku(0).Value Then
            RV_TorKb = 3
        Else
            RV_TorKb = 9
        End If
    End If
    RV_Kojyo = ""
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        If KB = 0 Then
            iB2_Torcd = RV_TorcdK
            iB2_Ukeir = RV_Ukeir
            B2lb_Torcd = RV_TorRName
        Else
            iB3_NTorcd = RV_TorcdK
            iB3_Ukeir = RV_Ukeir
            B3lb_NTorcd = RV_TorRName
        End If
    End If
End Sub
'+-----------------------------------+
'+       å_ñÒêÊåüçı(ìæà”êÊÅEî[ì¸êÊ)    +
'+-----------------------------------+
Private Sub Look_KNTorcd(KB As Integer)
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    If KB = 0 And B2Op_Hinmoku(0).Value Then
        RV_TorKb = 2
    End If
    If KB = 0 And B2Op_Hinmoku(1).Value Then
        RV_TorKb = 5
    End If
    If KB = 1 And B2Op_Hinmoku(0).Value Then
        RV_TorKb = 3
    End If
    If KB = 1 And B2Op_Hinmoku(1).Value Then
        RV_TorKb = 8
    End If
    If KB = 2 And B2Op_Hinmoku(1).Value Then
        RV_TorKb = 8
    End If
'-------------------------------------------' 2009.10.20 add start
    If KB = 3 Then
        RV_TorKb = 3
    End If
'-------------------------------------------' 2009.10.20 add end
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        Select Case KB
           Case 0
                iB3_KTorcd = RV_TorcdK
                B3lb_KTorcd = RV_TorRName
           Case 1
                iB3_NTorcd = RV_TorcdK
                B3lb_NTorcd = RV_TorRName
           Case 2
                iB3_DaikoK = RV_TorcdK
                B3lb_DaikoK = RV_TorRName
'-------------------------------------------' 2009.10.20 add start
           Case 3
                iB1_DaiCD = RV_TorcdK
                B1lb_DaiNM = RV_TorRName
'-------------------------------------------' 2009.10.20 add end
        End Select
    End If
End Sub
'+----------------------------+
'+       ñºèÃÉ}ÉXÉ^åüçı        +
'+    KB : 0     çHíˆ         +
'+       : 1     éüçHíˆ       +
'+       : 3     çÏã∆Å@Å@      +
'+       : 4     é˚óeäÌ(äÓñ{) +
'+       : 5     é˚óeäÌ1      +
'+       : 6     é˚óeäÌ2      +
'+       : 7     ì`ï[éÌóﬁî≠íç +
'+       : 8     ì`ï[éÌóﬁéÛíç +
'+       : 9     ïiî‘ãÊï™     +
'+       :13     ï‘ãpèÍèä     + ' 2009.10.21 add
'+       : 0à»äO íSìñ         +
'+----------------------------+
Private Sub Look_Kubun(KB As Integer)

    RV_Left = 0
    RV_Top = 0

    Select Case KB
        Case 0: RVI_Kubun = "çHíˆä«óù"
        Case 1: RVI_Kubun = "çHíˆä«óù"
        Case 3: RVI_Kubun = "çÏã∆ä«óù"
        Case 4, 5, 6: RVI_Kubun = "é˚óeäÌä«óù"
        Case 7: RVI_Kubun = "ì`ï[éÌóﬁî≠íç"
        Case 8: RVI_Kubun = "ì`ï[éÌóﬁéÛíç"
        Case 9: RVI_Kubun = "ïiî‘ãÊï™"
        Case 10: RVI_Kubun = "ã@äB"                     '2009/07/22
        Case 11: RVI_Kubun = "ã@äBNO"                   '2009/07/22
        Case 12: RVI_Kubun = "çHíˆä«óù"                 '2009/07/22
        Case 13: RVI_Kubun = "ï‘ãpèÍèä"     ' 2009.10.21 add
        Case Else: RVI_Kubun = "íSìñä«óù"
    End Select
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
        Select Case KB
            '''Case 0
            '''    iB1_Kotei = RV_Code
            '''    Call iB1_Kotei_LostFocus
            '''Case 1
            '''    iB1_NKotei = RV_Code
            '''    Call iB1_NKotei_LostFocus
            Case 2
                iB1_Tanto = RV_Code
                Call iB1_Tanto_LostFocus
'============================= *** 2009/07/22 Start
'            Case 3
'                iB3_Sagyo = RV_Code
'                Call iB3_Sagyo_LostFocus
'============================= *** 2009/07/22 End
            Case 4
                iB1_KUtsuwa = RV_Code
            Case 5
                iB3_Youki1 = RV_Code
            Case 6
                iB3_Youki2 = RV_Code
            Case 7
                iB3_DenSyu = RV_Code
                B3lb_DenSyu = RV_Kubun
            Case 8
                iB3_DenSyuJ = RV_Code
                B3lb_DenSyuJ = RV_Kubun
            Case 9
                iB1_Hinku = RV_Code
                B1lb_Hinku = RV_Kubun
'---------------------------------------------------------------------------' 2009.10.19 del start
''''            '=================================== *** 2009/07/22 Start
''''            Case 10
''''                iB3_Kikai = RV_Code
''''                Call iB3_Kikai_LostFocus
''''            Case 11
''''                iB3_KikaiNo = RV_Code
''''                Call iB3_KikaiNo_LostFocus
''''            Case 12
''''                iB3_Kotei = RV_Code
''''                Call iB3_Kotei_LostFocus
''''            '=================================== *** 2009/07/22 End
'---------------------------------------------------------------------------' 2009.10.19 del end
'---------------------------------------------------------------------------' 2009.10.21 add start
            Case 13
                iB3_HenCD = RV_Code
                B3lb_HenNM = RV_Kubun
'---------------------------------------------------------------------------' 2009.10.21 add end
        End Select
    End If

End Sub
'=========================================== *** 2009/07/22 Start
''+---------------------------------------------------------- 2004/12/13 Insert
''+----------------------------+
''+       ã@äBåüçı             +
''+----------------------------+
'Private Sub Look_Kikai()
'
'    RV_Left = 0
'    RV_Top = 0
'
'    CKP0610.Show vbModal
'    Unload CKP0610
'    Set CKP0610 = Nothing
'
'    If RV_Rtn Then
'       iB3_Kikai = RV_Code
'       Call iB3_Kikai_LostFocus
'    End If
'
'End Sub
''+---------------------------------------------------------- 2004/12/13 Insert End
'=========================================== *** 2009/07/22 End
'+----------------------------------------------------+
'+       É|ÉPÉbÉgáÇÅ@åüçı   2005/09/30í«â¡             +
'+----------------------------------------------------+
Private Sub Look_PNO()

    RV_Left = 0
    RV_Top = 0
    
    RV_Torcd = RTrim(iB2_Torcd)
    If B2Op_Hinmoku(0).Value Then
       RV_TorKb = 0
    Else
       RV_TorKb = 1
    End If
        
    CXM1035.Show vbModal
    Unload CXM1035
    Set CXM1035 = Nothing

    If RV_Rtn Then
       iB3_PNo = CisFun.RSetFld(RV_Code, 2, "#0")
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
'+  iB1_KUtsuwa
'+-----------------------------+
Private Sub iB1_KUtsuwa_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_KUtsuwa_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  B1Comb_Zaiko
'+-----------------------------+
Private Sub B1Comb_Zaiko_GotFocus()
    B1Comb_Zaiko.BackColor = gIGotSel
End Sub
Private Sub B1Comb_Zaiko_LostFocus()
    B1Comb_Zaiko.BackColor = gILostSel
End Sub
Private Sub B1Comb_Zaiko_Click()
    CisComboZaiko.CodeGet
    gStr = CisComboZaiko.Code
    Select Case gStr
        Case "2"    ' åWêî
            iB1_Zaiko.Visible = True
            iB1_ZaikoDay.Visible = False
            iB1_ZaikoL.Visible = False
            Zaiko_Label.Left = iB1_Zaiko.Left + iB1_Zaiko.Width + 60
            Zaiko_Label = "Åì"
            Zaiko_Label.Visible = True
        Case "3"    ' î≠íçì_
            iB1_Zaiko.Visible = False
            iB1_ZaikoDay.Visible = False
            iB1_ZaikoL.Visible = True
        Case "4"    ' ì˙êî
            iB1_Zaiko.Visible = False
            iB1_ZaikoDay.Visible = True
            iB1_ZaikoL.Visible = False
            Zaiko_Label.Left = iB1_ZaikoDay.Left + iB1_ZaikoDay.Width + 60
            Zaiko_Label = "ì˙"
            Zaiko_Label.Visible = True
        Case Else
            iB1_Zaiko.Visible = False
            iB1_ZaikoDay.Visible = False
            iB1_ZaikoL.Visible = False
            Zaiko_Label.Visible = False
    End Select
End Sub
'+-----------------------------+
'+  iB1_Hinku
'+-----------------------------+
Private Sub iB1_Hinku_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Hinku_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B1lb_Hinku = ""
    If MeisyoGet("ïiî‘ãÊï™", iB1_Hinku) Then
        B1lb_Hinku = SYM_Meisyo
    End If
End Sub
'+-----------------------------+
'+  B1Chk_JSKbn
'+-----------------------------+
Private Sub B1Chk_JSKbn_Click()
    If B1Chk_JSKbn.Value = 1 Then
       B1Chk_JSKbn.BackColor = mGotColor
       B1Chk_JSKbn.Caption = "Ç†ÇË"
    Else
       B1Chk_JSKbn.BackColor = mLostColor
       B1Chk_JSKbn.Caption = "Ç»Çµ"
    End If
End Sub
'+-----------------------------+
'+  iB1_Kotei
'+-----------------------------+
'''Private Sub iB1_Kotei_GotFocus()
'''    PB_Look.Tag = ActiveControl.Name
'''    PB_Look.Visible = True
'''End Sub
'''Private Sub iB1_Kotei_LostFocus()
'''    If ActiveControl.Name <> "PB_Look" Then
'''        PB_Look.Tag = ""
'''        PB_Look.Visible = False
'''    End If
'''    B1lb_Kotei = ""
'''    If MeisyoGet("çHíˆä«óù", iB1_Kotei) Then
'''        B1lb_Kotei = SYM_Meisyo
'''    End If
'''End Sub
'+-----------------------------+
'+  iB1_NKotei
'+-----------------------------+
'''Private Sub iB1_NKotei_GotFocus()
'''    PB_Look.Tag = ActiveControl.Name
'''    PB_Look.Visible = True
'''End Sub
'''Private Sub iB1_NKotei_LostFocus()
'''    If ActiveControl.Name <> "PB_Look" Then
'''        PB_Look.Tag = ""
'''        PB_Look.Visible = False
'''    End If
'''    B1lb_NKotei = ""
'''    If MeisyoGet("çHíˆä«óù", iB1_NKotei) Then
'''        B1lb_NKotei = SYM_Meisyo
'''    End If
'''End Sub
'+-----------------------------+
'+  iB1_Tanto
'+-----------------------------+
Private Sub iB1_Tanto_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Tanto_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B1lb_Tanto = ""
    If MeisyoGet("íSìñä«óù", iB1_Tanto) Then
        B1lb_Tanto = SYM_Meisyo
    End If
End Sub
'+-----------------------------+
'+  iB1_KataTor
'+-----------------------------+
Private Sub iB1_KataTor_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_KataTor_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B1lb_KataTor = ""
    If TorNmGet(iB1_KataTor, 9) Then
        B1lb_KataTor = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  iB1_Edit
'+-----------------------------+
Private Sub iB1_Edit_LostFocus()
    B1lb_EHinbn = EditHinbn(iH1_Hinbn, iB1_Edit)
End Sub
'+-----------------------------+
'+  iB1_Image
'+-----------------------------+
Private Sub iB1_Image_LostFocus()
'    If RTrim(Mid(iB1_Image, 1, 2)) <> "\\" Then
'        With CisFun
'            .FE_FullPath = iB1_Image
'            .GetUNCPath
'            iB1_Image = .FE_UNCPath
'        End With
'    End If
End Sub
'+-----------------------------+
'+  iB2_Torcd
'+-----------------------------+
Private Sub iB2_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB2_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B2lb_Torcd = ""
    If B2Op_Hinmoku(0).Value Then gInt = 3 Else gInt = 9
    If TorNmGet(iB2_Torcd, gInt) Then
        B2lb_Torcd = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  iB2_Ukeir
'+-----------------------------+
Private Sub iB2_Ukeir_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB2_Ukeir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  B2Op_Hinmoku
'+-----------------------------+
Private Sub B2Op_Hinmoku_Click(Index As Integer)
    If Index = 0 Then
       B2Op_Hinmoku(0).BackColor = mGotColor
       B2Op_Hinmoku(1).BackColor = mLostColor
       Back_Seban.Caption = "ìæîwî‘"
    Else
       B2Op_Hinmoku(0).BackColor = mLostColor
       B2Op_Hinmoku(1).BackColor = mGotColor
       Back_Seban.Caption = "îwî‘çÜ"
    End If
    B2Op_Hinmoku(0).Tag = Format(Index, "0")

    For Each gObj In Back_Juchu
        gObj.Visible = B2Op_Hinmoku(0).Value
    Next gObj
    For Each gObj In Back_Hachu
        gObj.Visible = B2Op_Hinmoku(1).Value
    Next gObj
'    B3_Area2.Visible = B2Op_Hinmoku(1).Value               2009/07/22
    
'-----------------------------------------------' 2009.10.20 add start
    If Index = 0 Then
       Call SyanaiTehai(False)  ' é–ì‡å¿íËçÄñ⁄Çè¡ãéÅAà íuïúãA
    Else
       Back_HachuPHA.Visible = True
    End If
'-----------------------------------------------' 2009.10.20 add end

    ' 2004/04/15 Mnt
    If FormAct Then
       Call iB2_Torcd_LostFocus
    End If
End Sub
'+-----------------------------+
'+  B3Op_SK
'+-----------------------------+
Private Sub B3Op_SK_Click(Index As Integer)
    If Index = 0 Then
       B3Op_SK(0).BackColor = mGotColor
       B3Op_SK(1).BackColor = mLostColor
       Back_Hachu(8).Enabled = False
       B3Chk_Naiji.Value = 0
    Else
       B3Op_SK(0).BackColor = mLostColor
       B3Op_SK(1).BackColor = mGotColor
       Back_Hachu(8).Enabled = True
    End If
End Sub
'+-----------------------------+
'+  iB3_Youki1
'+-----------------------------+
Private Sub iB3_Youki1_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_Youki1_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB3_HYmd
'+-----------------------------+
Private Sub iB3_HYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_HYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB3_Youki2
'+-----------------------------+
Private Sub iB3_Youki2_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_Youki2_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB3_KTorcd
'+-----------------------------+
Private Sub iB3_KTorcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_KTorcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B3lb_KTorcd = ""
    If B2Op_Hinmoku(0).Value Then gInt = 2 Else gInt = 5
    If TorNmGet(iB3_KTorcd, gInt) Then
        B3lb_KTorcd = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  B3Chk_Auto
'+-----------------------------+
Private Sub B3Chk_Auto_Click()
    If B3Chk_Auto.Value = 0 Then
       B3Chk_Auto.Caption = "éËìÆ"
    Else
       B3Chk_Auto.Caption = "é©ìÆ"
    End If
End Sub
'+-----------------------------+
'+  iB3_YSYmd
'+-----------------------------+
Private Sub iB3_YSYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_YSYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB3_YEYmd
'+-----------------------------+
Private Sub iB3_YEYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_YEYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  ChkB3_Daiko         2007/12/14 ADD
'+-----------------------------+
Private Sub ChkB3_Daiko_Click()
    If ChkB3_Daiko.Value = 1 Then
       ChkB3_Daiko.Caption = "Åyë„çsÅz"
       iB3_DaikoK.Enabled = True
    Else
       ChkB3_Daiko.Caption = "ÅyíºëóÅz"
       iB3_DaikoK.Enabled = False
       iB3_DaikoK = "": B3lb_DaikoK = ""
    End If
End Sub
'+-----------------------------+
'+  iB3_DaikoK          2007/12/14 ADD
'+-----------------------------+
Private Sub iB3_DaikoK_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_DaikoK_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B3lb_DaikoK = ""
    If B2Op_Hinmoku(0).Value Then gInt = 3 Else gInt = 8
    If TorNmGet(iB3_DaikoK, gInt) Then
        B3lb_DaikoK = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  iB3_NTorcd
'+-----------------------------+
Private Sub iB3_NTorcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_NTorcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B3lb_NTorcd = ""
    If B2Op_Hinmoku(0).Value Then gInt = 3 Else gInt = 8
    If TorNmGet(iB3_NTorcd, gInt) Then
        B3lb_NTorcd = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  iB3_Ukeir
'+-----------------------------+
Private Sub iB3_Ukeir_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_Ukeir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  B3Comb_Tehai
'+-----------------------------+
Private Sub B3Comb_Tehai_GotFocus()
    B3Comb_Tehai.BackColor = gIGotSel
End Sub
Private Sub B3Comb_Tehai_LostFocus()
    B3Comb_Tehai.BackColor = gILostSel
End Sub
Private Sub B3Comb_Tehai_Click()
    CisComboTehai.CodeGet
    gStr = CisComboTehai.Code
    
    '====================================================== *** 2008/06/09 Start
    If B2Op_Hinmoku(0).Value = True Then
       PB_UTN.Visible = True: PB_KTN.Visible = False: PB_STN.Visible = False
       Exit Sub
    End If
    '====================================================== *** 2008/06/09 End
    
    Select Case gStr
        Case "1"        ' íºëó
            iB3_NTorcd.Enabled = True
            iB3_Ukeir.Enabled = True
            Back_Hachu(4).Visible = True
            Back_Hachu(9).Visible = False
            Back_Hachu(10).Visible = False
            iB3_DaikoK = "": ChkB3_Daiko.Value = 0: B3lb_DaikoK = ""
            '-----* 2008/06/09 Start
            Back_Hachu(22).Visible = False
            Back_Hachu(5).Visible = True: Back_Hachu(14).Visible = True: Back_Hachu(15).Visible = True
            Back_Hachu(16).Visible = True: Back_Hachu(11).Visible = True: Back_Hachu(12).Visible = True
            Back_Hachu(23).Visible = False
            Back_Hachu(32).Visible = True       '2009/01/19
            
            
            PB_UTN.Visible = False: PB_KTN.Visible = True: PB_STN.Visible = True
            '-----* 2008/06/09 End
        Case "3"        ' ë„çså_ñÒêÊ                    2007/12/14 ADD
            iB3_NTorcd.Enabled = True
            iB3_Ukeir.Enabled = False
            Back_Hachu(4).Visible = False
            Back_Hachu(9).Visible = True
            Back_Hachu(10).Visible = False
            iB3_NTorcd = "": iB3_Ukeir = "": B3lb_NTorcd = ""
            '-----* 2008/06/09 Start
            Back_Hachu(22).Visible = False
            Back_Hachu(5).Visible = True: Back_Hachu(14).Visible = True: Back_Hachu(15).Visible = True
            Back_Hachu(16).Visible = True: Back_Hachu(11).Visible = True: Back_Hachu(12).Visible = True
            Back_Hachu(23).Visible = False
            Back_Hachu(32).Visible = False       '2009/01/19
            
            
            PB_UTN.Visible = False: PB_KTN.Visible = True: PB_STN.Visible = False
            '-----* 2008/06/09 End
        '-----* 2008/06/09 Start
        Case "4"
            iB3_NTorcd.Enabled = False
            iB3_Ukeir.Enabled = False
            iB3_NTorcd = "": iB3_Ukeir = "": B3lb_NTorcd = ""
            iB3_DaikoK = "": ChkB3_Daiko.Value = 0: B3lb_DaikoK = ""
            Back_Hachu(10).Visible = False
            Back_Hachu(5).Visible = False: Back_Hachu(14).Visible = False: Back_Hachu(15).Visible = False
            Back_Hachu(16).Visible = False: Back_Hachu(11).Visible = False: Back_Hachu(12).Visible = False
            Back_Hachu(4).Visible = False
            Back_Hachu(9).Visible = False
            Back_Hachu(22).Visible = True
            Back_Hachu(23).Visible = True
            
            Back_Hachu(32).Visible = False       '2009/01/19
            
           
           PB_UTN.Visible = False: PB_KTN.Visible = False: PB_STN.Visible = True
        '-----* 2008/06/09 End
        Case Else
            iB3_NTorcd.Enabled = False
            iB3_Ukeir.Enabled = False
            iB3_NTorcd = "": iB3_Ukeir = "": B3lb_NTorcd = ""
            iB3_DaikoK = "": ChkB3_Daiko.Value = 0: B3lb_DaikoK = ""
            Back_Hachu(4).Visible = False
            Back_Hachu(9).Visible = False
            Back_Hachu(10).Visible = True
            '-----* 2008/06/09 Start
            Back_Hachu(22).Visible = False
            Back_Hachu(5).Visible = True: Back_Hachu(14).Visible = True: Back_Hachu(15).Visible = True
            Back_Hachu(16).Visible = True: Back_Hachu(11).Visible = True: Back_Hachu(12).Visible = True
            Back_Hachu(23).Visible = False
            
            Back_Hachu(32).Visible = False       '2009/01/19
            
            
            PB_UTN.Visible = False: PB_KTN.Visible = True: PB_STN.Visible = False
            '-----* 2008/06/09 End
            
    End Select
End Sub
'+-----------------------------+
'+  B3Op_Heijun
'+-----------------------------+
Private Sub B3Op_Heijun_Click(Index As Integer)
    
    B3Op_Heijun(0).Tag = Index
    
    For Each gObj In B3Op_Heijun
        If gObj.Index = Index Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
    If Index = 0 Then
       iB3_HKeisu.Enabled = False
    Else
       iB3_HKeisu.Enabled = True
    End If
End Sub
'+-----------------------------+
'+  B3Chk_Week
'+-----------------------------+
Private Sub B3Chk_Week_Click(Index As Integer)
    If B3Chk_Week(Index).Value = 1 Then
       B3Chk_Week(Index).BackColor = mGotColor
    Else
       B3Chk_Week(Index).BackColor = mLostColor
    End If
End Sub
'+-----------------------------+
'+  B3Chk_Naiji
'+-----------------------------+
Private Sub B3Chk_Naiji_Click()
    If B3Chk_Naiji.Value = 1 Then
       B3Chk_Naiji.BackColor = mGotColor
       B3Chk_Naiji.Caption = "Ç∑ÇÈ"
    Else
       B3Chk_Naiji.BackColor = mLostColor
       B3Chk_Naiji.Caption = "ÇµÇ»Ç¢"
    End If
End Sub
'+-----------------------------+
'+  iB3_DenSyu(éÛíç)
'+-----------------------------+
Private Sub iB3_DenSyuJ_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_DenSyuJ_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    B3lb_DenSyuJ = ""
    If MeisyoGet("ì`ï[éÌóﬁéÛíç", iB3_DenSyuJ) Then
        B3lb_DenSyuJ = SYM_Meisyo
    End If
End Sub
'+-----------------------------+
'+  iB3_DenSyu(î≠íç)
'+-----------------------------+
Private Sub iB3_DenSyu_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_DenSyu_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    B3lb_DenSyu = ""
    If MeisyoGet("ì`ï[éÌóﬁî≠íç", iB3_DenSyu) Then
        B3lb_DenSyu = SYM_Meisyo
    End If
    If iB3_DenSyu = "3" Or iB3_DenSyu = "4" Then
        B3Chk_PrtkbS.Value = 1: Call B3Chk_PrtkbS_Click
        B3Chk_PrtkbJ.Value = 1: Call B3Chk_PrtkbJ_Click
    End If
End Sub
'''2007/12/11DEL'+----------------------------------+
'''2007/12/11DEL'+  B3Comb_NouhinJÅ@î[ïièëå`ë‘(éÛíç)
'''2007/12/11DEL'+----------------------------------+
'''2007/12/11DELPrivate Sub B3Comb_NouhinJ_GotFocus()
'''2007/12/11DEL    B3Comb_NouhinJ.BackColor = gIGotSel
'''2007/12/11DELEnd Sub
'''2007/12/11DELPrivate Sub B3Comb_NouhinJ_LostFocus()
'''2007/12/11DEL    B3Comb_NouhinJ.BackColor = gILostSel
'''2007/12/11DELEnd Sub
'''2007/12/11DEL'+----------------------------------+
'''2007/12/11DEL'+  B3Comb_NouhinÅ@î[ïièëå`ë‘(î≠íç)
'''2007/12/11DEL'+----------------------------------+
'''2007/12/11DELPrivate Sub B3Comb_Nouhin_GotFocus()
'''2007/12/11DEL    B3Comb_Nouhin.BackColor = gIGotSel
'''2007/12/11DELEnd Sub
'''2007/12/11DELPrivate Sub B3Comb_Nouhin_LostFocus()
'''2007/12/11DEL    B3Comb_Nouhin.BackColor = gILostSel
'''2007/12/11DELEnd Sub
'''2007/12/11DEL'+-----------------------------+
'''2007/12/11DEL'+  B1Comb_JyuryoÅ@éÛóÃèëå`ë‘
'''2007/12/11DEL'+-----------------------------+
'''2007/12/11DELPrivate Sub B3Comb_Jyuryo_GotFocus()
'''2007/12/11DEL    B3Comb_Jyuryo.BackColor = gIGotSel
'''2007/12/11DELEnd Sub
'''2007/12/11DELPrivate Sub B3Comb_Jyuryo_LostFocus()
'''2007/12/11DEL    B3Comb_Jyuryo.BackColor = gILostSel
'''2007/12/11DELEnd Sub
'+-----------------------------+
'+  B3Chk_PrtkbNÅ@î[ïièëî≠çsãÊï™
'+-----------------------------+
Private Sub B3Chk_PrtkbN_Click()
    If B3Chk_PrtkbN.Value = 0 Then
       B3Chk_PrtkbN.Caption = "î≠çsóL"
    Else
       B3Chk_PrtkbN.Caption = "î≠çsñ≥"
    End If
End Sub
'+-----------------------------+
'+  B3Chk_PrtkbSÅ@éxããèëî≠çsãÊï™
'+-----------------------------+
Private Sub B3Chk_PrtkbS_Click()
    If B3Chk_PrtkbS.Value = 0 Then
       B3Chk_PrtkbS.Caption = "î≠çsóL"
    Else
       B3Chk_PrtkbS.Caption = "î≠çsñ≥"
    End If
End Sub
'+-----------------------------+
'+  B3Chk_PrtkbJÅ@éÛóÃèëî≠çsãÊï™
'+-----------------------------+
Private Sub B3Chk_PrtkbJ_Click()
    If B3Chk_PrtkbJ.Value = 0 Then
       B3Chk_PrtkbJ.Caption = "î≠çsóL"
    Else
       B3Chk_PrtkbJ.Caption = "î≠çsñ≥"
    End If
End Sub
'-----* 2008/06/09 Start
'+-----------------------------+
'+  B3Chk_PrtkbSNÅ@éËîzãÊï™ÅFéxãã î[ïièëî≠çsãÊï™
'+-----------------------------+
Private Sub B3Chk_PrtkbSN_Click()
    If B3Chk_PrtkbSN.Value = 0 Then
       B3Chk_PrtkbSN.Caption = "î≠çsóL"
    Else
       B3Chk_PrtkbSN.Caption = "î≠çsñ≥"
    End If
End Sub
'-----* 2008/06/09 End
'======================== 2009/01/19 Start
'+-----------------------------+
'+  B3Chk_PrtkbSN1Å@éËîzãÊï™ÅFíºëó éxããì`ï[î≠çsãÊï™
'+-----------------------------+
Private Sub B3Chk_PrtkbSN1_Click()
    If B3Chk_PrtkbSN1.Value = 0 Then
       B3Chk_PrtkbSN1.Caption = "î≠çsóL"
    Else
       B3Chk_PrtkbSN1.Caption = "î≠çsñ≥"
    End If
End Sub
'======================== 2009/01/19 End

'''2007/12/11DEL'+------------------------------------+
'''2007/12/11DEL'+  B1Comb_JyuryoJÅ@éÛóÃèëå`ë‘(éÛíç)
'''2007/12/11DEL'+------------------------------------+
'''2007/12/11DELPrivate Sub B3Comb_JyuryoJ_GotFocus()
'''2007/12/11DEL    B3Comb_JyuryoJ.BackColor = gIGotSel
'''2007/12/11DELEnd Sub
'''2007/12/11DELPrivate Sub B3Comb_JyuryoJ_LostFocus()
'''2007/12/11DEL    B3Comb_JyuryoJ.BackColor = gILostSel
'''2007/12/11DELEnd Sub
'+-----------------------------+
'+  B3Chk_JuryoÅ@éÛóÃèëî≠çs
'+-----------------------------+
Private Sub B3Chk_Juryo_Click()
    If B3Chk_Juryo.Value = 0 Then
       B3Chk_Juryo.Caption = "î≠íçéû"
    Else
       B3Chk_Juryo.Caption = "éÛóÃéû"
    End If
End Sub
'+-----------------------------+
'+  B3Chk_PrtkbNJÅ@î[ïièëî≠çsãÊï™
'+-----------------------------+
Private Sub B3Chk_PrtkbNJ_Click()
    If B3Chk_PrtkbNJ.Value = 0 Then
       B3Chk_PrtkbNJ.Caption = "î≠çsóL"
    Else
       B3Chk_PrtkbNJ.Caption = "î≠çsñ≥"
    End If
End Sub
'+-----------------------------+
'+  B3Chk_PrtkbSJÅ@éxããèëî≠çsãÊï™
'+-----------------------------+
Private Sub B3Chk_PrtkbSJ_Click()
    If B3Chk_PrtkbSJ.Value = 0 Then
       B3Chk_PrtkbSJ.Caption = "î≠çsóL"
    Else
       B3Chk_PrtkbSJ.Caption = "î≠çsñ≥"
    End If
End Sub
'+-----------------------------+
'+  B3Chk_PrtkbJJÅ@éÛóÃèëî≠çsãÊï™
'+-----------------------------+
Private Sub B3Chk_PrtkbJJ_Click()
    If B3Chk_PrtkbJJ.Value = 0 Then
       B3Chk_PrtkbJJ.Caption = "î≠çsóL"
    Else
       B3Chk_PrtkbJJ.Caption = "î≠çsñ≥"
    End If
End Sub
'+-----------------------------+
'+  B3Chk_JuryoJÅ@éÛóÃèëî≠çs(éÛíç)
'+-----------------------------+
Private Sub B3Chk_JuryoJ_Click()
    If B3Chk_JuryoJ.Value = 0 Then
       B3Chk_JuryoJ.Caption = "éÛíçéû"
    Else
       B3Chk_JuryoJ.Caption = "èoâ◊éû"
    End If
End Sub
'======================================= *** 2009/07/22 Start
'+-----------------------------+
'+  B3Chk_SijisyoÅ@éwé¶èëî≠çs(éÛíç)
'+-----------------------------+
Private Sub B3Chk_Sijisyo_Click()
    If B3Chk_Sijisyo.Value = 0 Then
       B3Chk_Sijisyo.Caption = "î≠çsóL"
    Else
       B3Chk_Sijisyo.Caption = "î≠çsñ≥"
    End If
End Sub
'======================================= *** 2009/07/22 End
'+-----------------------------+
'+  B3Comb_MeaisaiÅ@ñæç◊èëî≠çs
'+-----------------------------+
Private Sub B3Comb_Meaisai_GotFocus()
    B3Comb_Meaisai.BackColor = gIGotSel
End Sub
Private Sub B3Comb_Meaisai_LostFocus()
    B3Comb_Meaisai.BackColor = gILostSel
End Sub
'+-----------------------------+
'+  iB3_PNo
'+-----------------------------+
Private Sub iB3_PNo_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_PNo_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
'+--------------------+
'+     B1Op_Buhin     +
'+--------------------+
Private Sub B1Op_Buhin_Click(Index As Integer)
    B1Op_Buhin(0).Tag = Index
End Sub
'+---------------------+
'+     B3Chk_Gouho     +
'+---------------------+
Private Sub B3Chk_Gouho_Click()
    If B3Chk_Gouho.Value = 0 Then
       B3Chk_Gouho.Caption = "çÜå˚"
    Else
       B3Chk_Gouho.Caption = "ï‚ãã"
    End If
End Sub

'+----------------------+
'+     B3Chk_GouhoJ     +
'+----------------------+
Private Sub B3Chk_GouhoJ_Click()
    If B3Chk_GouhoJ.Value = 0 Then
       B3Chk_GouhoJ.Caption = "çÜå˚"
    Else
       B3Chk_GouhoJ.Caption = "ï‚ãã"
    End If
End Sub

'+----------------------+
'+     B1Chk_MaeHak     +
'+----------------------+
Private Sub B3Chk_Maehak_Click()
    If B3Chk_MaeHak.Value = 0 Then
       B3Chk_MaeHak.Caption = "ï\é¶ñ≥"
    Else
       B3Chk_MaeHak.Caption = "ï\é¶óL"
    End If
End Sub
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
'============================================== *** 2009/07/22 Start
'+-----------------------------+
'+  B3Op_Heijun
'+-----------------------------+
Private Sub B3Op_Hinku_Click(Index As Integer)
    
    B3Op_HinKu(0).Tag = Index
    
    For Each gObj In B3Op_HinKu
        If gObj.Index = Index Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
'    If INDEX = 0 Then
'       iB3_HKeisu.Enabled = False
'    Else
'       iB3_HKeisu.Enabled = True
'    End If
End Sub
'============================================== *** 2009/07/22 End


'**************************************
'*                                    *
'*      èàóùãÊï™Ç…ÇÊÇËãÊï™ñºï\é¶        *
'*                                    *
'**************************************
'+-----------------------+
'+      äÓñ{èÓïÒ
'+-----------------------+
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
'+-----------------------+
'+      éÊà¯êÊèÓïÒ
'+-----------------------+
Private Sub SyoriName2(SYKB As String)
   Dim W_Color     As Long
   W_Color = &H80000008 'èâä˙êFæØƒ(çï)
   Select Case SyoriKB2
     Case "ADD"
          SyoriNM2 = "í«â¡"
          K_Sykbnm2.ForeColor = &HFFFF&   '(â©êF)
          PB_ADD.Visible = False
     Case "MNT"
          SyoriNM2 = "èCê≥"
          K_Sykbnm2.ForeColor = &HFF0000  '(ê¬êF)
     Case "DEL"
          SyoriNM2 = "çÌèú"
          K_Sykbnm2.ForeColor = &HFF&     '(ê‘êF)
          W_Color = &HFF&
     Case "LOK"
          SyoriNM = "è∆âÔ"
          K_Sykbnm2.ForeColor = &HFF00FF   '(ÉsÉìÉN)
          PB_ADD.Visible = False
          PB_MNT.Visible = False
          PB_DEL.Visible = False
     Case Else
          SyoriNM2 = ""
   End Select
   For Each gObj In Me.Controls
       If gObj.Name Like "iB3*" Then gObj.ForeColor = W_Color
       If gObj.Name Like "B3Comb*" Then gObj.ForeColor = W_Color
'      * PF-Key ï\é¶ê›íË *
       If SyoriKB <> "LOC" Then
          If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
             gObj.Name = "PB_DEL" Then
             gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB2)
          End If
       End If
   Next gObj
   Timer3.Enabled = False
   If SyoriNM2 = "" Then
      K_Sykbnm2.Visible = False
      K_Sykbnm2.Caption = ""
      Exit Sub
   End If
   K_Sykbnm2.Caption = SyoriNM2
   K_Sykbnm2.Visible = True
   Timer3.Enabled = True
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
'*******************************
'*      èàóùãÊï™ñºì_ñ≈êßå‰      *
'*******************************
Private Sub Timer3_Timer()
   If K_Sykbnm2.Caption <> "" Then
      K_Sykbnm2.Caption = ""
   Else
      K_Sykbnm2.Caption = SyoriNM2
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
    Dim B3Mode      As Boolean
    Dim B3Color     As Long
    H1Mode = False
    H1Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
    B2Mode = False
    B2Color = gPLostSel
    B3Mode = False
    B3Color = gPLostSel
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
          PB_Chg.Tag = "TOR"
          PB_Chg.BtnCaption = "éÊà¯êÊ"
       Case "B2"
          B2Mode = True
          B2Color = gPGotSel
          PB_Chg.Tag = "BASE"
          PB_Chg.BtnCaption = "äÓñ{"
       Case "B3"
          B3Mode = True
          B3Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B1_Area2.Enabled = B1Mode Or B2Mode Or B3Mode
    B1_Area2.BackColor = B1Color
    
    B2_Area1.Enabled = B2Mode
    B2_Area2.BackColor = B2Color
    
    B3_Area1.Enabled = B3Mode
    B3_Area1.BackColor = B3Color
    B3_Area2.Enabled = B3Mode
    B3_Area2.BackColor = B3Color
    
    PB_Look.Visible = H1Mode
    PB_ADD.Visible = H1Mode
    PB_MNT.Visible = H1Mode
    PB_DEL.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/03/31 DELETE START
'    If Rv_Call3 <> "" Then
'        PB_CAN.Visible = Not B2Mode
'    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/03/31 DELETE END
    
    PB_Chg.Visible = B1Mode Or B2Mode
    
    PB_Zairyo.Visible = B1Mode
    PB_Image.Visible = B1Mode
    PB_ImageDisp.Visible = B1Mode
    
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
    B3_Area3.Visible = B3Mode
'    PB_UTN.Visible = B3Mode And B2Op_Hinmoku(0).Value
'    PB_KTN.Visible = B3Mode And B2Op_Hinmoku(1).Value
'    PB_STN.Visible = B3Mode And B2Op_Hinmoku(1).Value
'    PB_UTN.Enabled = B3Mode And B2Op_Hinmoku(0).Value
'    PB_KTN.Enabled = B3Mode And B2Op_Hinmoku(1).Value
'    PB_STN.Enabled = B3Mode And B2Op_Hinmoku(1).Value
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
    
    PB_Kotei.Visible = B3Mode                               '209/07/24
    
    
' ¿ﬁ–∞çÄñ⁄(ì¸óÕã÷é~)
    Dummy.Enabled = False
'
End Function
'**************************************
'*      Head/Body èâä˙âªèàóù           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String
    RowColFlg = True
    If ClrType = "ALL" Then wbc_Enm = HeadBody & "*" Else wbc_Enm = Trim(ClrType)
    On Error GoTo HeadBodyClear_ER
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then
                .Text = ""
              '  .Enabled = True
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
                .BackColor = mLostColor
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B1" Or HeadBody = "B" Then
        CisVsGrid.Clear
        CisVsGrid2.Clear
'---------------------------------------------------' 2009.10.20 del start
''''        CisVsGrid3.Clear
''''        CisVsGrid4.Clear
'---------------------------------------------------' 2009.10.20 del end
        'Call CisVsGridS.Clear                                      2009/07/22
        '+----------------------------------- 2004/12/13 Insert
        'Call CisVsGridK.Clear                                      2009/07/22
        '+----------------------------------- 2004/12/13 Insert End
    End If
    If HeadBody = "B3" Then
        'Call CisVsGridS.EditEnable("C")                            2009/07/22
        'Call CisVsGridS.Clear                                      2009/07/22
        '+----------------------------------- 2004/12/13 Insert
        'Call CisVsGridK.EditEnable("C")                            2009/07/22
        'Call CisVsGridK.Clear                                      2009/07/22
        '+----------------------------------- 2004/12/13 Insert End
        CisVsGrid2.Clear
'---------------------------------------------------' 2009.10.20 del start
''''        CisVsGrid3.Clear
''''        CisVsGrid4.Clear
'---------------------------------------------------' 2009.10.20 del send
        Call B3Comb_Tehai_Click                 '2007/12/14 ADD
        B3Op_HinKu(0).Value = True              '2009/07/22
    End If
'---------------------------------------------------' 2009.10.16 del start
''''    If HeadBody = "B4" Then
''''        iB3_SetHinbn = "": iB3_TSu = ""
''''        CisVsGrid4.Clear
''''    End If
'---------------------------------------------------' 2009.10.16 del end
    Image1.Picture = Nothing                        ' 2009.10.20 add : É|ÉCÉìÉgÇÃÉCÉÅÅ[ÉWÇè¡ãé
    
    RowColFlg = False
    Exit Sub
HeadBodyClear_ER:
    RowColFlg = False
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
        If Trim(vsGrid1.TextMatrix(1, 1)) <> "" Then
           vsGrid1.SetFocus
        End If
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then
            GoTo ReturnPress_Ed
        End If
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B2" Then
        If Not Body2Chk Then GoTo ReturnPress_Ed
' << çÌèúèàóù >>
        If SyoriKB2 = "DEL" Then
            If Not DBDeleteTor Then
               RowColFlg = True
               Call CisVsGrid.vsColor(è¡ãéëSïî)
               RowColFlg = False
               GoTo ReturnPress_Ed
            End If
            Call HeadBodyClear("B3")
            Call DispChange(ProcHB)
            Call SyoriName2(SyoriKB2)
            GoTo ReturnPress_Ed
        Else
            GridBack = ProcHB
            ProcHB = "B3"
            Call DispChange(ProcHB)
             ' ∏ﬁÿØƒﬁì‡óeï\é¶å„ÅA√∑Ωƒì¸óÕ
            'Call CisVsGridS.EditVisible("C")                           2009/07/22
            '+-------------------------------------- 2004/12/13 Insert
            'Call CisVsGridK.EditVisible("C")                           2009/07/22
            '+-------------------------------------- 2004/12/13 Insert End
        End If
        iB3_Seban.SetFocus
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B3" Then
        If Not Body3Chk Then
'           Call CisVsGridS.EditVisible("C")
           GoTo ReturnPress_Ed
        End If
        
        RowColFlg = True
        Call CisVsGrid.vsColor(è¡ãéëSïî)
        RowColFlg = False
        
        If Not DBPut(False) Then
            GoTo ReturnPress_Ed
        End If
                
        If Not DBPutTor Then
'           Call CisVsGridS.EditVisible("C")
           GoTo ReturnPress_Ed
        End If
        
        ProcHB = GridBack
        Call DispChange(ProcHB)
        If GridBack = "B1" Then
           SyoriKB2 = ""
           Call SyoriName2(SyoriKB2)
           'Call CisVsGridS.EditEnable("C")                                 2009/07/22
           '+-------------------------------------- 2004/12/13 Insert
           'Call CisVsGridK.EditEnable("C")                                 2009/07/22
           '+-------------------------------------- 2004/12/13 Insert End
           Call DispChange(ProcHB)
           vsGrid1.SetFocus
        Else
           Call SyoriName2(SyoriKB2)
           iB2_Torcd.SetFocus
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
    Head1Chk = False
    
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

' É}ÉXÉ^ì¸óÕ SELECTï∂ æØƒ
    gSL_Select = " SELECT * "
    gSL_Select = gSL_Select & " FROM ïiî‘É}ÉXÉ^ "
    gSL_Select = gSL_Select & " WHERE ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    If DBInput("INV") Then
       If SyoriKB = "ADD" Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    ïiî‘É}ÉXÉ^ìoò^çœ            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          iH1_Hinbn.SetFocus
          Exit Function
       End If
    Else
       If SyoriKB <> "ADD" Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    ïiî‘É}ÉXÉ^ñ¢ìoò^            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          iH1_Hinbn.SetFocus
          Exit Function
       End If
    End If
    If SyoriKB = "DEL" Then
'       gSL_Select = "SELECT ïiî‘ FROM ïiî‘édì¸êÊÉ}ÉXÉ^ "
'       gSL_Select = gSL_Select & " WHERE édì¸êÊ = '" & Trim(iH1_Torcd) & "'"
'       If HNSRead(gSL_Select, 1) Then
'          ReDim gMBMsg(5)
'          gMBMsg(2) = "     ïiî‘édì¸êÊÉ}ÉXÉ^Ç…édì¸êÊÇ™ìoò^çœÇ›Ç≈Ç∑ÅB       "
'          gMBMsg(4) = "             çÌèúèoóàÇ‹ÇπÇÒÅB                     "
'          gMBTitle = "ÉGÉâÅ[ÅEçƒì¸óÕ"
'          gMBButton = "E"
'          gRtn = MBOX
'          Exit Function
'       End If
'       gSL_Select = "SELECT éxããéxï•êÊ FROM éÊà¯êÊÉ}ÉXÉ^ "
'       gSL_Select = gSL_Select & " Where éxããéxï•êÊ = '" & RTrim(iH1_Torcd) & "'"
'       If TRMRead(gSL_Select, 1) Then
'          ReDim gMBMsg(5)
'          gMBMsg(2) = "     édì¸êÊÉ}ÉXÉ^Ç…éxããéxï•êÊÇ∆ÇµÇƒìoò^çœÇ›Ç≈Ç∑ÅB       "
'          gMBMsg(4) = "             çÌèúèoóàÇ‹ÇπÇÒÅB                     "
'          gMBTitle = "ÉGÉâÅ[ÅEçƒì¸óÕ"
'          gMBButton = "E"
'          gRtn = MBOX
'          Exit Function
'       End If
    End If
    
    ' ïiî‘ï“èWèâä˙ï\é¶
    If SyoriKB = "ADD" Then
       iB1_Edit = HNM_Edit
        B1Op_Buhin(0).Value = True
    End If
    ' ïiî‘ï“èWï\é¶
    B1lb_EHinbn = EditHinbn(iH1_Hinbn, iB1_Edit)
       
    Head1Chk = True
End Function
'****************************
'*      Body1¡™Ø∏èàóù
'****************************
Private Function Body1Chk() As Boolean
    
    Body1Chk = False
    Set mBodyErrFld = Nothing
    
    ' ïiñº
    If Trim(iB1_HinNM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    ïiñºÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_HinNM.SetFocus
       Set mBodyErrFld = iB1_HinNM
       Exit Function
    End If
    ' îwî‘çÜîºäpÉ`ÉFÉbÉN
    If RTrim(RTrim(iB1_SSeban)) <> "" And Not HalfSizeChk(RTrim(iB1_SSeban)) Then
        With CisFun
             .MB_Lines = 4
             .MB_MSG(2) = "    é–ì‡îwî‘ÇÕîºäpÇ…Çƒì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
        End With
        iB1_SSeban.SetFocus
        Exit Function
    End If
    ' ç›å…ä«óù
    With CisComboZaiko
        .CodeGet
        gStr = .Code
    End With
    If gStr = "2" Then
       If CisFun.Val2(iB1_Zaiko) = 0 Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    ç›å…åWêîÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB1_Zaiko.SetFocus
           Set mBodyErrFld = iB1_Zaiko
           Exit Function
        End If
    End If
    If gStr = "3" Then
       If CisFun.Val2(iB1_ZaikoL) = 0 Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    ç›å…î≠íçì_Çì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB1_ZaikoL.SetFocus
           Set mBodyErrFld = iB1_ZaikoL
           Exit Function
        End If
    End If
    If gStr = "4" Then
       If CisFun.Val2(iB1_ZaikoDay) = 0 Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    ç›å…ì˙êîÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB1_ZaikoDay.SetFocus
           Set mBodyErrFld = iB1_ZaikoDay
           Exit Function
        End If
    End If
    ' ïiî‘ãÊï™
    B1lb_Hinku = ""
    If Trim(iB1_Hinku) <> "" Then
        If MeisyoGet("ïiî‘ãÊï™", iB1_Hinku) Then
            B1lb_Hinku = SYM_Meisyo
        Else
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ïiî‘ãÊï™ñ¢ìoò^            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iB1_Hinku.SetFocus
           Set mBodyErrFld = iB1_Hinku
           Exit Function
        End If
    End If
    ' çHíˆ
'''    B1lb_Kotei = ""
'''    If Trim(iB1_Kotei) <> "" Then
'''        If MeisyoGet("çHíˆä«óù", iB1_Kotei) Then
'''            B1lb_Kotei = SYM_Meisyo
'''        Else
'''            With CisFun
'''                .MB_Lines = 4
'''                .MB_MSG(2) = "    çHíˆñ¢ìoò^            "
'''                .MB_Title = ""
'''                .MB_Button = Error
'''                .MBOX
'''            End With
'''           iB1_Kotei.SetFocus
'''           Set mBodyErrFld = iB1_Kotei
'''           Exit Function
'''        End If
'''    End If
    ' éüçHíˆ
'''    B1lb_NKotei = ""
'''    If Trim(iB1_NKotei) <> "" Then
'''        If MeisyoGet("çHíˆä«óù", iB1_NKotei) Then
'''            B1lb_NKotei = SYM_Meisyo
'''        Else
'''            With CisFun
'''                .MB_Lines = 4
'''                .MB_MSG(2) = "    éüçHíˆñ¢ìoò^            "
'''                .MB_Title = ""
'''                .MB_Button = Error
'''                .MBOX
'''            End With
'''           iB1_NKotei.SetFocus
'''           Set mBodyErrFld = iB1_NKotei
'''           Exit Function
'''        End If
'''    End If
    ' íSìñ
    B1lb_Tanto = ""
    If Trim(iB1_Tanto) <> "" Then
        If MeisyoGet("íSìñä«óù", iB1_Tanto) Then
            B1lb_Tanto = SYM_Meisyo
        Else
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    íSìññ¢ìoò^            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iB1_Tanto.SetFocus
           Set mBodyErrFld = iB1_Tanto
           Exit Function
        End If
    End If
'-----------------------------------------------------------' 2009.10.21 add start
    ' ë„ï\ìæà”êÊ
    B1lb_DaiNM = ""
    If Trim(iB1_DaiCD.Text) <> "" Then
       If TorNmGet(iB1_DaiCD, 3) Then
          B1lb_DaiNM = TRM_RName
       Else
          With CisFun
             .MB_Lines = 4
             .MB_MSG(2) = "    ë„ï\ìæà”êÊ  ñ¢ìoò^              "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
          End With
          iB1_DaiCD.SetFocus
          Exit Function
       End If
    End If
'-----------------------------------------------------------' 2009.10.21 add end

'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2006/12/19 Insert Start
    If BD1Chk_Flg <> 0 Then GoTo Body1Chk_ED
'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2006/12/19 Insert End
    
    ' ≤“∞ºﬁÃß≤Ÿñº
    If Trim(iB1_Image) <> "" Then
'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2006/12/19 Update Start
'       If InStr(1, iB1_Image, "\") <> 0 Or _
'          InStr(1, iB1_Image, ":") <> 0 Then
       If InStr(1, iB1_Image, ":") <> 0 Then
'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2006/12/19 Update End
          With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    ÉCÉÅÅ[ÉWÉtÉ@ÉCÉãñºÇ…ÇÕÉtÉ@ÉCÉãñºÇì¸óÕÇµÇƒâ∫Ç≥Ç¢ÅB            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
          End With
          iB1_Image.SetFocus
          Set mBodyErrFld = iB1_Image
          Exit Function
       End If
       If Dir(mImage & iB1_Image) = "" Then
          With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    ÉCÉÅÅ[ÉWÉtÉ@ÉCÉãÇ™ìoò^Ç≥ÇÍÇƒÇ¢Ç»Ç¢Ç© ÉtÉ@ÉCÉãñºÇ™ïsê≥Ç≈Ç∑ÅB            "
               .MB_MSG(4) = "                       ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH             "
               .MB_Title = ""
               .MB_Button = OK_CAN
               If Not .MBOX Then
                  iB1_Image.SetFocus
                  Set mBodyErrFld = iB1_Image
                  Exit Function
               End If
           End With
       End If
    End If

'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2006/12/19 Insert Start
Body1Chk_ED:
'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2006/12/19 Insert End
    Body1Chk = True
End Function
'****************************
'*      BODY2¡™Ø∏èàóù
'****************************
Private Function Body2Chk() As Boolean
    Dim ReadFlg     As Boolean
    Body2Chk = False
    
    If Trim(iB2_Torcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    éÊà¯êÊÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB2_Torcd.SetFocus
       Exit Function
    End If
' éÊà¯êÊÉ`ÉFÉbÉN
    gSL_Select = " SELECT * "
    gSL_Select = gSL_Select & " FROM éÊà¯êÊÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " WHERE éÊà¯êÊãÊï™ = " & B2Op_Hinmoku(0).Tag & ""
    gSL_Select = gSL_Select & "   AND éÊà¯êÊCD = '" & RTrim(iB2_Torcd) & "'"
    If Not TRMRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    éÊà¯êÊÉ}ÉXÉ^ñ¢ìoò^            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB2_Torcd.SetFocus
       Exit Function
    End If
    
'    If B2Op_Hinmoku(0).Value Then
'        If TRM.éÊà¯êÊãÊï™ <> 0 Then
'            With CisFun
'                .MB_Lines = 4
'                .MB_MSG(2) = "    ïiñ⁄(éÛíç)éwíËéûéÊà¯êÊÇÕéÛíçÇéwíËÇµÇƒâ∫Ç≥Ç¢            "
'                .MB_Title = ""
'                .MB_Button = Error
'                .MBOX
'            End With
'           iB2_Torcd.SetFocus
'           Exit Function
'        End If
'    Else
'        If TRM.éÊà¯êÊãÊï™ <> 1 Then
'            With CisFun
'                .MB_Lines = 4
'                .MB_MSG(2) = "    ïiñ⁄(î≠íç)éwíËéûéÊà¯êÊÇÕî≠íçÇéwíËÇµÇƒâ∫Ç≥Ç¢            "
'                .MB_Title = ""
'                .MB_Button = Error
'                .MBOX
'            End With
'           iB2_Torcd.SetFocus
'           Exit Function
'        End If
'    End If
' éÛì¸
    If TRM.é–ì‡ãÊï™ = 0 And Trim(iB2_Ukeir) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   éÛì¸Çì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB2_Ukeir.SetFocus
       Exit Function
    End If
' éÛì¸ë∂ç›É`ÉFÉbÉN
    If Trim(iB2_Ukeir) <> "" Then
        gSL_Select = " SELECT * "
        gSL_Select = gSL_Select & " FROM éÊà¯êÊéÛì¸É}ÉXÉ^ "
        gSL_Select = gSL_Select & " WHERE éÊà¯êÊãÊï™ = " & B2Op_Hinmoku(0).Tag & ""
        gSL_Select = gSL_Select & "   AND éÊà¯êÊCD = '" & RTrim(iB2_Torcd) & "'"
        gSL_Select = gSL_Select & "   And éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
        If Not TUMRead(gSL_Select) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    éÊà¯êÊÇ…éÛì¸Ç™ìoò^Ç≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÅB            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iB2_Ukeir.SetFocus
           Exit Function
        End If
    End If
' É}ÉXÉ^ì¸óÕ SELECTï∂ æØƒ
    With vsGrid1
        ReadFlg = False
        If B2Op_Hinmoku(0).Value Then
           gStr = "0"
        Else
           gStr = "1"
        End If
        iB2_Torcd.Tag = ""
        For mIx = 1 To .Rows - 1
            If Trim(.TextMatrix(mIx, SCol(39))) = "" Then Exit For
            If RTrim(iB2_Torcd) = Trim(.TextMatrix(mIx, SCol(39))) And _
               RTrim(iB2_Ukeir) = Trim(.TextMatrix(mIx, SCol(2))) Then
               If RTrim(gStr) <> Trim(.TextMatrix(mIx, SCol(40))) Then
                  With CisFun
                     .MB_Lines = 4
                     .MB_MSG(2) = "    ïiñ⁄à·Ç¢Ç…Çƒìoò^çœ            "
                     .MB_Title = ""
                     .MB_Button = Error
                     .MBOX
                  End With
                  iB2_Ukeir.SetFocus
                  Exit Function
               End If
            End If
            If RTrim(iB2_Torcd) = Trim(.TextMatrix(mIx, SCol(39))) And _
               RTrim(iB2_Ukeir) = Trim(.TextMatrix(mIx, SCol(2))) And _
               RTrim(gStr) = Trim(.TextMatrix(mIx, SCol(40))) Then
               ReadFlg = True
               iB2_Torcd.Tag = mIx
               .Row = mIx
               Exit For
            End If
        Next mIx
        
        If ReadFlg Then
           If SyoriKB2 = "ADD" Then
              With CisFun
                   .MB_Lines = 4
                   .MB_MSG(2) = "    ïiî‘éÊà¯êÊÉ}ÉXÉ^ìoò^çœ            "
                   .MB_Title = ""
                   .MB_Button = Error
                   .MBOX
              End With
              iB2_Torcd.SetFocus
              Exit Function
           End If
           ' éÊà¯êÊèÓïÒï\é¶
           Call HeadBodyClear("B3")
           Call CisComboTehai.ComboSet
'''2007/12/11DEL           Call CisComboHNou.ComboSet
'''2007/12/11DEL           Call CisComboHJyu.ComboSet
           Call CisComboHMei.ComboSet
'''2007/12/11DEL           Call CisComboJNou.ComboSet
'''2007/12/11DEL           Call CisComboJJyu.ComboSet                   '2004/04/08
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
           Call CisComboJsya.ComboSet
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
           
           RowColFlg = True
           Call CisVsGrid.vsColor(å¬ï )
           RowColFlg = False
           
           Call GridToText
       Else
           If SyoriKB2 <> "ADD" Then
              With CisFun
                   .MB_Lines = 4
                   .MB_MSG(2) = "    ïiî‘éÊà¯êÊÉ}ÉXÉ^ñ¢ìoò^            "
                   .MB_Title = ""
                   .MB_Button = Error
                   .MBOX
              End With
              iB2_Torcd.SetFocus
              Exit Function
           End If
       End If
    
    End With
    
    ' äÓñ{èÓïÒÇÊÇËèâä˙ílï\é¶
    If SyoriKB2 = "ADD" Then
       Call HeadBodyClear("B3")
       Call CisComboTehai.ComboSet
'''2007/12/11DEL       Call CisComboHNou.ComboSet
'''2007/12/11DEL       Call CisComboHJyu.ComboSet
       Call CisComboHMei.ComboSet
'''2007/12/11DEL       Call CisComboJNou.ComboSet
'''2007/12/11DEL       Call CisComboJJyu.ComboSet
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
       Call CisComboJsya.ComboSet
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
       
       iB3_Syuyo1 = CisFun.RSetFld(CisFun.Val2(iB1_KSyuyo), 0, iB3_Syuyo1.cFormat)
       iB3_Youki1 = iB1_KUtsuwa
       B3Op_SK(0).Value = True
       iB3_KTorcd = TRM.å_ñÒêÊ
'-------------------------------------------------------- 2006/02/23 Insert
       If TRM.å_ñÒêÊãÊï™ = 1 Then
          iB3_KTorcd = TRM.éÊà¯êÊCD
       End If
'-------------------------------------------------------- 2006/02/23 Insert End
       B3Op_Heijun(TRM.ïΩèÄâªãÊï™).Value = True
       iB3_HKeisu = CisFun.RSetFld(TRM.ïΩèÄâªåWêî, 0, iB3_HKeisu.cFormat)
       If TRM.é–ì‡ãÊï™ = 1 Then
          B3lb_Tehai.Visible = True
          With CisComboTehai
               .Code = "2"
               .TextGet
               B3lb_Tehai = .CombText
          End With
          B3Comb_Tehai.Visible = False
'-------------------------------------------------------' 2009.10.20 add start
          Call SyanaiTehai(True)    ' é–ì‡éËîzçÄñ⁄Çï\é¶Ç∑ÇÈ
'-------------------------------------------------------' 2009.10.20 add end
       Else
          B3lb_Tehai.Visible = False
          B3Comb_Tehai.Visible = True
'-------------------------------------------------------' 2009.10.20 add start
          Call SyanaiTehai(False)   ' é–ì‡éËîzçÄñ⁄Çè¡ãéÇ∑ÇÈ
'-------------------------------------------------------' 2009.10.20 add end
       End If
'*--------------------------------------------------<< 2004/06/29 Insert Start >>
          B3Comb_Tehai.Tag = TRM.é–ì‡ãÊï™
'*--------------------------------------------------<< 2004/06/29 Insert End >>
       If Trim(TUM.ÉTÉCÉNÉã) <> "" Then
          iB3_Cycle1 = Mid(TUM.ÉTÉCÉNÉã, 1, 1)
          iB3_Cycle2 = Mid(TUM.ÉTÉCÉNÉã, 2, 2)
          iB3_Cycle3 = Mid(TUM.ÉTÉCÉNÉã, 4, 2)
       End If
       If B2Op_Hinmoku(0).Value Then
          iB3_DenSyuJ = TRM.ì`ï[éÌóﬁ
'''2007/12/11DEL           ' î[ïièëå`ë‘
'''2007/12/11DEL           With CisComboJNou
'''2007/12/11DEL                .Code = TRM.î[ïièëå`ë‘
'''2007/12/11DEL                .TextGet
'''2007/12/11DEL           End With
           Call iB3_DenSyuJ_LostFocus
'''2007/12/11DEL           ' éÛóÃèëå`ë‘
'''2007/12/11DEL           With CisComboJJyu
'''2007/12/11DEL                .Code = TRM.éÛóÃèëå`ë‘
'''2007/12/11DEL                .TextGet
'''2007/12/11DEL           End With
           B3Chk_PrtkbNJ.Value = TRM.î[ïièëî≠çsãÊï™
           B3Chk_PrtkbSJ.Value = TRM.éxããèëî≠çsãÊï™
           B3Chk_PrtkbJJ.Value = TRM.éÛóÃèëî≠çsãÊï™
           B3Chk_JuryoJ.Value = TRM.éÛóÃèëî≠çs
           
           '====================================== *** 2009/07/06 Start
           '''iB3_Syukaba = TRM.èoâ◊èÍ
           '====================================== *** 2009/07/06 End

           '====================================== *** 2009/07/22 Start
           B3Chk_Sijisyo = TRM.éwé¶èëî≠çsãÊï™
           '====================================== *** 2009/07/22 End

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
           If TRM.çÜï‚ãÊï™ = "H" Then
              B3Chk_GouhoJ.Value = 1
           End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
           
       Else
          iB3_DenSyu = TRM.ì`ï[éÌóﬁ
'''2007/12/11DEL           ' î[ïièëå`ë‘
'''2007/12/11DEL           With CisComboHNou
'''2007/12/11DEL                .Code = TRM.î[ïièëå`ë‘
'''2007/12/11DEL                .TextGet
'''2007/12/11DEL           End With
'''2007/12/11DEL           ' éÛóÃèëå`ë‘
'''2007/12/11DEL           With CisComboHJyu
'''2007/12/11DEL                .Code = TRM.éÛóÃèëå`ë‘
'''2007/12/11DEL                .TextGet
'''2007/12/11DEL           End With
           ' ñæç◊èëî≠çs
           With CisComboHMei
                .Code = TRM.ñæç◊èëî≠çs
                .TextGet
           End With
           
           B3Chk_PrtkbN.Value = TRM.î[ïièëî≠çsãÊï™
           B3Chk_PrtkbS.Value = TRM.éxããèëî≠çsãÊï™
           B3Chk_PrtkbJ.Value = TRM.éÛóÃèëî≠çsãÊï™
           B3Chk_Juryo.Value = TRM.éÛóÃèëî≠çs
           Call iB3_DenSyu_LostFocus
           
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
           If TRM.çÜï‚ãÊï™ = "H" Then
              B3Chk_Gouho.Value = 1
           End If
           B3Chk_MaeHak.Value = TRM.ëOçHíˆãÊï™
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
'-------------------------------------------------------' 2009.10.20 add start
           Call SyanaiTehai(True)    ' é–ì‡éËîzçÄñ⁄Çï\é¶Ç∑ÇÈ
'-------------------------------------------------------' 2009.10.20 add end
           
       End If
       Call iB3_KTorcd_LostFocus
    End If
    
    Body2Chk = True
End Function
'**************************************
'*      BODY3¡™Ø∏èàóù(éÊà¯êÊèÓïÒ)
'**************************************
Private Function Body3Chk() As Boolean
    Dim ReadFlg     As Boolean
    Dim wSel        As String           ' 2009.10.21 add
    Body3Chk = False

    'Call CisVsGridS.EditEnable                                         2009/07/22
    '+----------------------------------------- 2004/12/13 Insert
    'Call CisVsGridK.EditEnable                                         2009/07/22
    '+----------------------------------------- 2004/12/13 Insert End
'    Call CisVsGridS.EditVisible
'    vsGridS.Row = 1
'    Call CisVsGridS.EditEnable("C")

    iB3_Seban.SetFocus
    ' îwî‘çÜîºäpÉ`ÉFÉbÉN
    If RTrim(RTrim(iB3_Seban)) <> "" And Not HalfSizeChk(RTrim(iB3_Seban)) Then
        With CisFun
             .MB_Lines = 4
             .MB_MSG(2) = "    " & Back_Seban & "ÇÕîºäpÇ…Çƒì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
        End With
        iB3_Seban.SetFocus
        Exit Function
    End If
 ' Ç©ÇÒÇŒÇÒéûÅAïKê{ì¸óÕ
    If B3Op_SK(0).Value Then
       ' îwî‘çÜ
       If Trim(iB3_Seban) = "" Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    îwî‘çÜÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iB3_Seban.SetFocus
          Exit Function
       End If
    ' ÉTÉCÉNÉã
       If Trim(iB3_Cycle1) = "" And Trim(iB3_Cycle2) = "" And Trim(iB3_Cycle3) = "" Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    ÉTÉCÉNÉãÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iB3_Cycle1.SetFocus
          Exit Function
       End If
    ' é˚óeêî
       If CisFun.Val2(iB3_Syuyo1) = 0 Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    é˚óeêîá@Çì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iB3_Syuyo1.SetFocus
          Exit Function
       End If
    End If
    
    ' ÉTÉCÉNÉã
    If Trim(iB3_Cycle1) <> "" Or Trim(iB3_Cycle2) <> "" Or Trim(iB3_Cycle3) <> "" Then
       If Trim(iB3_Cycle1) = "" Or Trim(iB3_Cycle2) = "" Or Trim(iB3_Cycle3) = "" Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    ÉTÉCÉNÉãÇê≥ÇµÇ≠ì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iB3_Cycle1.SetFocus
          Exit Function
       End If
    End If
    ' ïœçXì˙
    If Trim(iB3_HYmd) <> "" Then
       If Not iB3_HYmd.cChkResult Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    ïœçXì˙Çê≥ÇµÇ≠ì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iB3_HYmd.SetFocus
          Exit Function
       End If
       ' ïœçXï÷
       If Trim(iB3_HBin) = "" Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    ïœçXï÷Çì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iB3_HBin.SetFocus
          Exit Function
       End If
       ' é˚óeêîáA
       If CisFun.Val2(iB3_Syuyo2) = 0 Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    é˚óeêîáAÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iB3_Syuyo2.SetFocus
          Exit Function
       End If
    End If
    ' é˚óeêîá@
    If CisFun.Val2(iB3_Syuyo1) = 0 And CisFun.Val2(iB3_Syuyo2) <> 0 Then
       With CisFun
           .MB_Lines = 4
           .MB_MSG(2) = "    é˚óeêîá@Çì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
           .MB_Title = ""
           .MB_Button = Error
           .MBOX
       End With
       iB3_Syuyo1.SetFocus
       Exit Function
    End If
    ' ïœçXì˙
    If CisFun.Val2(iB3_Syuyo2) <> 0 And Trim(iB3_HYmd) = "" Then
       With CisFun
           .MB_Lines = 4
           .MB_MSG(2) = "    ïœçXì˙Çì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
           .MB_Title = ""
           .MB_Button = Error
           .MBOX
       End With
       iB3_HYmd.SetFocus
       Exit Function
    End If
    ' å_ñÒêÊ
    If Trim(iB3_KTorcd) = "" Then
       With CisFun
           .MB_Lines = 4
           .MB_MSG(2) = "    å_ñÒêÊÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
           .MB_Title = ""
           .MB_Button = Error
           .MBOX
       End With
       iB3_KTorcd.SetFocus
       Exit Function
    End If
    B3lb_KTorcd = ""
    If B2Op_Hinmoku(0).Value Then gInt = 2 Else gInt = 5
    If TorNmGet(iB3_KTorcd, gInt) Then
        B3lb_KTorcd = TRM_RName
    Else
       With CisFun
           .MB_Lines = 4
           .MB_MSG(2) = "    å_ñÒêÊñ¢ìoò^            "
           .MB_Title = ""
           .MB_Button = Error
           .MBOX
       End With
       iB3_KTorcd.SetFocus
       Exit Function
    End If

'-----     -----     -----     -----     -----     -----     -----     -----     < 2006/02/16 DEL START >
'
''    ' ÉAÉhÉåÉX
''    If CisFun.Val2(iB3_Adress) <> 0 Then
''        gSL_Select = "Select * from ïiî‘çﬁóøÉrÉÖÅ[ "
''        gSL_Select = gSL_Select & " Where ÉAÉhÉåÉX = '" & iB3_Adress & "'"
''        If ZTMRead(gSL_Select, 1) Then
''            If ZTM.ãÊï™ = 0 Then
''                If ZTM.ïiî‘ <> iH1_Hinbn Or _
''                   ZTM.éÊà¯êÊ <> iB2_Torcd Or _
''                   ZTM.éÛì¸ <> iB2_Ukeir Then
''
''                    With CisFun
''                        .MB_Lines = 6
''                        .MB_MSG(2) = "    ìØàÍÇÃÉAÉhÉåÉXÇ™ë∂ç›ÇµÇ‹Ç∑ÅB               "
''                        .MB_MSG(4) = "    ïi î‘  ÅF" & ZTM.ï\é¶ïiî‘
''                        .MB_MSG(5) = "    éÊà¯êÊÅF" & ZTM.éÊà¯êÊ
''                        .MB_MSG(6) = "    éÛ ì¸  ÅF" & ZTM.éÛì¸
''                        .MB_Title = ""
''                        .MB_Button = Error
''                        .MBOX
''                        iB3_Adress.SetFocus
''                        Exit Function
''                    End With
''                End If
''            Else
''                With CisFun
''                    .MB_Lines = 8
''                    .MB_MSG(2) = "    ìØàÍÇÃÉAÉhÉåÉXÇ™ë∂ç›ÇµÇ‹Ç∑ÅB               "
''                    .MB_MSG(4) = "    çﬁ éø ÅF" & ZTM.çﬁéø
''                    If ZTM.ê°ñ@ãÊï™ = 0 Then
''                        .MB_MSG(5) = "    ê° ñ@ ÅF" & Format(ZTM.î¬å˙, "###0.00") & " x " & _
''                                                    Format(ZTM.ïù, "###0.00") & " x " & _
''                                                    Format(ZTM.í∑Ç≥, "###0.00")
''                    End If
''                    .MB_MSG(6) = "    éËîzêÊ ÅF" & ZTM.éÊà¯êÊ
''                    .MB_MSG(7) = "    éÛ ì¸  ÅF" & ZTM.éÛì¸
''                    .MB_Title = ""
''                    .MB_Button = Error
''                    .MBOX
''                End With
''                iB3_Adress.SetFocus
''                Exit Function
''            End If
''        End If
''    Else
''        iB3_Adress = ""
''    End If
'
'-----     -----     -----     -----     -----     -----     -----     -----     < 2006/02/16 DEL  END  >
    
    ' î≠íçî‰ó¶
    Dim wHiritsu        As Long
    gLong = CisFun.Val2(iB3_Hiri)
    wHiritsu = 0
    If gLong <> 0 Then
       If gLong > 100 Then
          With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    î≠íçî‰ó¶ÇÕÇPÇOÇOà»â∫Ç…Çƒê›íËÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
          End With
          iB3_Hiri.SetFocus
          Exit Function
       End If
       wHiritsu = gLong
    End If
    If Not B2Op_Hinmoku(0).Value Then
        With vsGrid1
             gInt = 0
             For mIx = 1 To .Rows - 1
                 If Trim(.TextMatrix(mIx, SCol(1))) = "" Then Exit For
                 If Trim(.TextMatrix(mIx, SCol(40))) = "1" Then
                    gInt = gInt + 1
                    If Trim(.TextMatrix(mIx, SCol(39))) <> Trim(iB2_Torcd) Or _
                       Trim(.TextMatrix(mIx, SCol(2))) <> Trim(iB2_Ukeir) Then
                       wHiritsu = wHiritsu + CisFun.Val2(Replace(.TextMatrix(mIx, SCol(6)), "%", ""))
                    End If
                 End If
             Next mIx
             If gInt > 1 Then
                 If wHiritsu <> 100 Then
                    With CisFun
                         .MB_Lines = 5
                         .MB_MSG(2) = "    î≠íçî‰ó¶ÇÃçáåvÇ™ÇPÇOÇOÇ≈ÇÕÇ†ÇËÇ‹ÇπÇÒ            "
                         .MB_MSG(4) = "    Å@Å@Å@Å@Å@ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH            "
                         .MB_Title = ""
                         .MB_Button = Yes_No
                         If Not .MBOX Then
                            iB3_Hiri.SetFocus
                            Exit Function
                         End If
                    End With
                 End If
             End If
        End With
    End If
    ' óLå¯ä˙å¿
    If Trim(iB3_YSYmd) <> "" Then
       If Not iB3_YSYmd.cChkResult Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    óLå¯ä˙å¿ÅiäJénÅjÇê≥ÇµÇ≠ì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iB3_YSYmd.SetFocus
          Exit Function
       End If
    End If
    If Trim(iB3_YEYmd) <> "" Then
       If Not iB3_YEYmd.cChkResult Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    óLå¯ä˙å¿ÅièIóπÅjÇê≥ÇµÇ≠ì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iB3_YEYmd.SetFocus
          Exit Function
       End If
    End If
    If Trim(iB3_YSYmd) <> "" And Trim(iB3_YEYmd) <> "" Then
       If iB3_YSYmd > iB3_YEYmd Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    óLå¯ä˙å¿Çê≥ÇµÇ≠ì¸óÕÇµÇƒâ∫Ç≥Ç¢ÅiäJénÅÑèIóπÅj            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iB3_YSYmd.SetFocus
          Exit Function
       End If
    End If
    ' éËîzãÊï™
    With CisComboTehai
        .CodeGet
        gStr = .Code
    End With
'*--------------------------------------------------<< 2004/06/29 Insert Start >>
    If B3Comb_Tehai.Tag = "0" Then
        If gStr = "2" Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    é–ì‡ÇÕëIëÇ≈Ç´Ç‹ÇπÇÒ            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           B3Comb_Tehai.SetFocus
           Exit Function
        End If
    End If
'*--------------------------------------------------<< 2004/06/29 Insert End >>
    ' î[ì¸êÊ
    If gStr = "1" Then
       If Trim(iB3_NTorcd) = "" Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    î[ì¸êÊÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB3_NTorcd.SetFocus
           Exit Function
        End If
        B3lb_NTorcd = ""
        If B2Op_Hinmoku(0).Value Then gInt = 3 Else gInt = 8
        If TorNmGet(iB3_NTorcd, gInt) Then
            B3lb_NTorcd = TRM_RName
        Else
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    î[ì¸êÊñ¢ìoò^            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB3_NTorcd.SetFocus
           Exit Function
        End If
        If Trim(iB3_Ukeir) = "" Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    î[ì¸êÊéÛì¸Çì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB3_Ukeir.SetFocus
           Exit Function
        End If
    ' éÛì¸ë∂ç›É`ÉFÉbÉN
        gSL_Select = " SELECT * "
        gSL_Select = gSL_Select & " FROM éÊà¯êÊéÛì¸É}ÉXÉ^ "
        gSL_Select = gSL_Select & " WHERE éÊà¯êÊãÊï™ = " & B2Op_Hinmoku(0).Tag & ""
        gSL_Select = gSL_Select & "   AND éÊà¯êÊCD = '" & RTrim(iB3_NTorcd) & "'"
        gSL_Select = gSL_Select & "   And éÛì¸ = '" & RTrim(iB3_Ukeir) & "'"
        If Not TUMRead(gSL_Select) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    î[ì¸êÊéÛì¸Ç™ìoò^Ç≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÅB            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iB3_Ukeir.SetFocus
           Exit Function
        End If
    End If
    
'-------------------------------------------------------------------------------' 2009.10.21 add start
'    If gStr = "2" Then
    ' É|ÉCÉìÉg
        Image1.Picture = Nothing
        If iB3_Point <> "" Then
           POM.É|ÉCÉìÉg = RTrim(iB3_Point.Text)
           If POMRead("") = True Then
              On Error Resume Next
              Set Image1.Picture = LoadPicture(gIniExe & "Image\" & RTrim(POM.ÉCÉÅÅ[ÉW))
              On Error GoTo 0
           Else
              With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    É|ÉCÉìÉg  ñ¢ìoò^                "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
              End With
              iB3_Point.SetFocus
              Exit Function
           End If
        End If
    ' ï‘ãpèÍèä
        B3lb_HenNM = ""
        If Trim(iB3_HenCD.Text) <> "" Then
           wSel = "SELECT * FROM ñºèÃÉ}ÉXÉ^"
           wSel = wSel & " WHERE ãÊï™ñºèÃ = 'ï‘ãpèÍèä'"
           wSel = wSel & "   AND íl       = '" & Trim(iB3_HenCD.Text) & "'"
           If SYMRead(wSel) = True Then
              B3lb_HenNM = SYM.ílñºèÃ
           Else
              With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ï‘ãpèÍèä  ñ¢ìoò^                "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
              End With
              iB3_HenCD.SetFocus
              Exit Function
           End If
        End If
    ' å„çHíˆ
        If CisFun.Val2(iB3_AtoCD.Text) = 0 Then
           iB3_AtoCD.Text = ""
        Else
           iB3_AtoCD.Text = Format(CisFun.Val2(iB3_AtoCD.Text), "000")
        End If
        B3lb_AtoNM = ""
        If Trim(iB3_AtoCD.Text) <> "" Then
           wSel = "SELECT TM.ó™èÃ, M2.ílñºèÃ ê≥éÆñºèÃ"                                     ' , tm.*, m1.*
           wSel = wSel & " FROM  éÊà¯êÊÉ}ÉXÉ^ TM"
           wSel = wSel & "       LEFT OUTER JOIN ñºèÃÉ}ÉXÉ^ M1"
           wSel = wSel & "              ON  M1.ãÊï™ñºèÃ = 'çHíˆÇ©ÇÒÇŒÇÒî≠çsÉpÉ^Å[Éì'"
           wSel = wSel & "              AND M1.íl       = '" & Trim(iB3_AtoCD.Text) & "'"  ' å„çHíˆ
           wSel = wSel & "              AND M1.ílñºèÃ   = '" & Trim(iB2_Torcd.Text) & "'"  ' éÊà¯êÊ
           wSel = wSel & "       LEFT OUTER JOIN ñºèÃÉ}ÉXÉ^ M2"
           wSel = wSel & "              ON  M2.ãÊï™ñºèÃ = 'ï‘ãpèÍèä'"
           wSel = wSel & "              AND M2.íl       = M1.âpêîéö3"                      ' ï‘ãpèÍèä1
           wSel = wSel & " WHERE TM.éÊà¯êÊCD   = M1.âpêîéö2"                               ' ïîèê1
           wSel = wSel & "   AND TM.éÊà¯êÊãÊï™ = 1"                                        ' éÊà¯êÊãÊï™ = î≠íç  ' 2009.10.23 add
'''''''''''wSel = wSel & "   AND TM.é–ì‡ãÊï™   = 1"                                        ' é–ì‡ãÊï™   = é–ì‡  ' 2009.10.23 del
           If TRMRead(wSel) = True Then
              B3lb_AtoNM = Trim(TRM.ó™èÃ) & " " & Trim(TRM.ê≥éÆñºèÃ)   ' ê≥éÆñºèÃÇ…ï‘ãpèÍèä1ÇÉZÉbÉgÇµÇƒÇ¢ÇÈ
           Else
              With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    å„çHíˆ  ñ¢ìoò^                  "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
              End With
              iB3_AtoCD.SetFocus
              Exit Function
           End If
        End If
'    End If
'-------------------------------------------------------------------------------' 2009.10.21 add end

'*--------------------------------------------------<< 2007/12/14 Insert Start >>
    ' ë„çså_ñÒêÊ
    If gStr = "3" Then
        If ChkB3_Daiko.Value <> 0 Then
            If Trim(iB3_DaikoK) = "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    ë„çså_ñÒêÊÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                iB3_DaikoK.SetFocus
                Exit Function
             End If
             B3lb_NTorcd = ""
             If B2Op_Hinmoku(0).Value Then gInt = 3 Else gInt = 8
             If TorNmGet(iB3_DaikoK, gInt) Then
                 B3lb_DaikoK = TRM_RName
             Else
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    ë„çså_ñÒêÊÅ@ñ¢ìoò^            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                iB3_DaikoK.SetFocus
                Exit Function
             End If
        End If
    End If
'*--------------------------------------------------<< 2007/12/14 Insert End >>
    
    ' ïΩèÄâª
    If B3Op_Heijun(1).Value Or B3Op_Heijun(2).Value Then
       If iB3_Cycle1 <> "1" Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    ïΩèÄâªÇ…Çƒì˙ñîÇÕï÷ÇéwíËéûÇÕÉTÉCÉNÉãÇÃÇPåÖñ⁄ÇÕ'1'à»äOê›íËèoóàÇ‹ÇπÇÒ                "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB3_Cycle1.SetFocus
           Exit Function
       End If
    End If
    ' ÉçÉbÉgá@
    If CisFun.Val2(iB3_LSu) <> 0 Then
       If CisFun.Val2(iB3_Syuyo1) <> 0 Then
          If CisFun.Val2(iB3_LSu) Mod CisFun.Val2(iB3_Syuyo1) <> 0 Then
             With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ÉçÉbÉgá@ÇÕé˚óeêîá@ÇÃî{êîÇ≈Ç»ÇØÇÍÇŒÇ»ÇËÇ‹ÇπÇÒÅB            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB3_LSu.SetFocus
            Exit Function
          End If
       End If
    End If
    ' ÉçÉbÉgáA
    If CisFun.Val2(iB3_LSu2) <> 0 Then
       If CisFun.Val2(iB3_Syuyo2) <> 0 Then
          If CisFun.Val2(iB3_LSu2) Mod CisFun.Val2(iB3_Syuyo2) <> 0 Then
             With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ÉçÉbÉgáAÇÕé˚óeêîáAÇÃî{êîÇ≈Ç»ÇØÇÍÇŒÇ»ÇËÇ‹ÇπÇÒÅB            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB3_LSu2.SetFocus
            Exit Function
          End If
       Else
         With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    ÉçÉbÉgáAÇÕé˚óeêîáAÇ™ñ¢ì¸óÕÇÃà◊ÅAì¸óÕèoóàÇ‹ÇπÇÒÅB            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
         End With
         iB3_LSu2.SetFocus
         Exit Function
       End If
    End If
    ' ójì˙
    mCHK = False
    For Each gObj In B3Chk_Week
        If gObj.Value = 1 Then
           mCHK = True
           Exit For
        End If
    Next gObj
    
    If mCHK Then
       If CisFun.Val2(iB3_Day(0)) <> 0 Or CisFun.Val2(iB3_Day(1)) <> 0 Or _
          CisFun.Val2(iB3_Day(2)) <> 0 Or CisFun.Val2(iB3_Day(3)) <> 0 Or _
          CisFun.Val2(iB3_Day(4)) <> 0 Then
         With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    ójì˙ñîÇÕì˙ÇÕÇ«ÇøÇÁÇ©àÍï˚ÇÃì¸óÕÇµÇ©èoóàÇ‹ÇπÇÒÅB            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
         End With
         B3Chk_Week(0).SetFocus
         Exit Function
       End If
    End If
    ' ì˙
    For gInt = 0 To 4
        If CisFun.Val2(iB3_Day(gInt)) > 31 Then
            With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    ê≥ÇµÇ¢ì˙Ç…Çƒì¸óÕÇµÇƒâ∫Ç≥Ç¢ÅiÇPÅ`ÇRÇPÅj            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
            End With
            iB3_Day(gInt).SetFocus
            Exit Function
        End If
    Next gInt
    
    Dim DayChk(4) As Long
    Erase DayChk
    gLong = 0
    For mIx = 0 To 4
        If CisFun.Val2(iB3_Day(mIx)) <> 0 Then
           DayChk(gLong) = CisFun.Val2(iB3_Day(mIx))
            For mIx2 = mIx + 1 To 4
                If Trim(iB3_Day(mIx2)) <> "" Then
                   If DayChk(gLong) = CisFun.Val2(iB3_Day(mIx2)) Then
                      With CisFun
                           .MB_Lines = 4
                           .MB_MSG(2) = "    ì˙Ç™èdï°ÇµÇƒÇ‹Ç∑ÅB            "
                           .MB_Title = ""
                           .MB_Button = Error
                           .MBOX
                      End With
                      iB3_Day(mIx2).SetFocus
                      Exit Function
                   End If
                End If
            Next mIx2
           gLong = gLong + 1
        End If
    Next mIx
    For mIx = 0 To 3
        If DayChk(mIx) = 0 Then Exit For
        For mIx2 = mIx + 1 To 4
            If DayChk(mIx2) <> 0 Then
               If DayChk(mIx) > DayChk(mIx2) Then
                  gLong = DayChk(mIx2)
                  DayChk(mIx2) = DayChk(mIx)
                  DayChk(mIx) = gLong
               End If
            End If
        Next mIx2
    Next mIx
    mIx2 = 0
    For mIx = 0 To 4
        iB3_Day(mIx) = ""
        If DayChk(mIx) <> 0 Then
           iB3_Day(mIx2) = CisFun.RSetFld(DayChk(mIx2), 0, iB3_Day(mIx2).cFormat)
           mIx2 = mIx2 + 1
        End If
    Next mIx
    ' ì`ï[éÌóﬁ
    If B2Op_Hinmoku(0).Value Then
'===Åy éÛíççÄñ⁄É`ÉFÉbÉN Åz============================================================================'
'-----     -----     -----     -----     -----     -----     -----     -----     < 2006/02/16 ADD START >
'
    '---(( ÉAÉhÉåÉX ))-------------------------------------------------------'
        If CisFun.Val2(iB3_Adress) <> 0 Then
            gSL_Select = "SELECT HN.ï\é¶ïiî‘, HT.éÊà¯êÊ, HT.éÛì¸"
            gSL_Select = gSL_Select & " FROM ïiî‘éÊà¯êÊÉ}ÉXÉ^ HT"
            gSL_Select = gSL_Select & " LEFT OUTER JOIN ïiî‘É}ÉXÉ^ HN"
            gSL_Select = gSL_Select & " ON   HN.ïiî‘ = HT.ïiî‘"
            gSL_Select = gSL_Select & " WHERE ht.ïiñ⁄ = 0"
            gSL_Select = gSL_Select & " AND   ht.ÉAÉhÉåÉX = '" & iB3_Adress & "'"
            gSL_Select = gSL_Select & " AND"
            gSL_Select = gSL_Select & " NOT"
            gSL_Select = gSL_Select & "     ( HT.ïiî‘ = '" & RTrim$(iH1_Hinbn) & "'"
            gSL_Select = gSL_Select & " AND   HT.éÊà¯êÊ = '" & RTrim$(iB2_Torcd) & "'"
            gSL_Select = gSL_Select & " AND   HT.éÛì¸ = '" & RTrim$(iB2_Ukeir) & "'"
            gSL_Select = gSL_Select & " )"
            If HTMRead(gSL_Select, 1) Then
                With CisFun
                    .MB_Lines = 6
                    .MB_MSG(2) = "    ìØàÍÇÃÉAÉhÉåÉXÇ™ë∂ç›ÇµÇ‹Ç∑ÅB               "
                    .MB_MSG(4) = "    ïi î‘  ÅF" & HTM.ï\é¶ïiî‘
                    .MB_MSG(5) = "    éÊà¯êÊÅF" & HTM.éÊà¯êÊ
                    .MB_MSG(6) = "    éÛ ì¸  ÅF" & HTM.éÛì¸
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                    iB3_Adress.SetFocus
                    Exit Function
                End With
            End If
        Else
            iB3_Adress = ""
        End If
'
'-----     -----     -----     -----     -----     -----     -----     -----     < 2006/02/16 ADD  END  >
       B3lb_DenSyuJ = "": SaveKeitai = ""
       If Trim(iB3_DenSyuJ) <> "" Then
          If MeisyoGet("ì`ï[éÌóﬁéÛíç", iB3_DenSyuJ) Then
             B3lb_DenSyuJ = SYM_Meisyo
             SaveKeitai = CisFun.Left2(SYM.âpêîéö1, 1)
          Else
             With CisFun
                  .MB_Lines = 4
                  .MB_MSG(2) = "    ì`ï[éÌóﬁñ¢ìoò^            "
                  .MB_Title = ""
                  .MB_Button = Error
                  .MBOX
             End With
             iB3_DenSyuJ.SetFocus
             Exit Function
          End If
'''2007/12/11DEL             CisComboJNou.CodeGet
'''2007/12/11DEL             CisComboJJyu.CodeGet
'''2007/12/11DEL             If RTrim(iB3_DenSyuJ) = "2" And (CisComboJNou.Code <> "2" Or CisComboJJyu.Code <> "2") Then
'''2007/12/11DEL                With CisFun
'''2007/12/11DEL                    .MB_Lines = 4
'''2007/12/11DEL                    .MB_MSG(1) = "    ì`ï[éÌóﬁÇ™ ÇiÇ`ÇlÇ`ÇÃéûÅAî[ïièëå`ë‘ÅEéÛóÃèëå`ë‘ÇÕÉtÉ@ÉCÉãèoóÕå¿íËÇ≈Ç∑ÅB            "
'''2007/12/11DEL                    .MB_MSG(3) = "    î[ïièëå`ë‘ÅEéÛóÃèëå`ë‘ÇÉtÉ@ÉCÉãèoóÕÇ…ê›íËÇµÇ‹Ç∑ÅB            "
'''2007/12/11DEL                    .MB_Title = ""
'''2007/12/11DEL                    .MB_Button = Error
'''2007/12/11DEL                    .MBOX
'''2007/12/11DEL                End With
'''2007/12/11DEL                CisComboJNou.Code = "2"
'''2007/12/11DEL                CisComboJNou.TextGet
'''2007/12/11DEL                CisComboJJyu.Code = "2"
'''2007/12/11DEL                CisComboJJyu.TextGet
'''2007/12/11DEL            End If
'''2007/12/11DEL            If B3Comb_NouhinJ.ListIndex = 0 And B3Comb_JyuryoJ.ListIndex = 0 Then
'''2007/12/11DEL                With CisFun
'''2007/12/11DEL                    .MB_Lines = 4
'''2007/12/11DEL                    .MB_MSG(2) = "    î[ïièëå`ë‘ÅEéÛóÃèëå`ë‘ÇÇê›íËÇµÇƒâ∫Ç≥Ç¢ÅB            "
'''2007/12/11DEL                    .MB_Title = ""
'''2007/12/11DEL                    .MB_Button = Error
'''2007/12/11DEL                    .MBOX
'''2007/12/11DEL                End With
'''2007/12/11DEL               B3Comb_NouhinJ.SetFocus
'''2007/12/11DEL               Exit Function
'''2007/12/11DEL            End If
        Else
'''2007/12/11DEL            B3Comb_NouhinJ.ListIndex = 0
'''2007/12/11DEL            B3Comb_JyuryoJ.ListIndex = 0
            B3Chk_PrtkbNJ.Value = 1
            B3Chk_PrtkbSJ.Value = 1
            B3Chk_PrtkbJJ.Value = 1
            B3Chk_JuryoJ.Value = 0
        End If
'============================================================= *** 2009/07/22 Start
'-----     -----     -----     -----     -----     -----     -----     -----     < 2006/02/16 MNT START >
'        If CisFun.Val2(iB3_JPNo) > maxPNo Then
'           With CisFun
'                .MB_Lines = 4
'                .MB_MSG(2) = "    PNO(É|ÉPÉbÉgNO)ÇÕ" & maxPNo & "à»â∫Ç…Çƒê›íËÇµÇƒâ∫Ç≥Ç¢            "
'                .MB_Title = ""
'                .MB_Button = Error
'                .MBOX
'           End With
'           iB3_JPNo.SetFocus
'           Exit Function
'        End If
'-----     -----     -----     -----     -----     -----     -----     -----     < 2006/02/16 MNT  END  >
'============================================================= *** 2009/07/22 End
    Else
'===Åy î≠íççÄñ⁄É`ÉFÉbÉN Åz============================================================================'
    
'-----     -----     -----     -----     -----     -----     -----     -----     < 2006/02/16 ADD START >
'
    '---(( ÉAÉhÉåÉX ))-------------------------------------------------------'
'                       << ïîïi >>
'
        If CisFun.Val2(iB3_Adress) <> 0 Then
            gSL_Select = "SELECT HN.ï\é¶ïiî‘, HT.éÊà¯êÊ, HT.éÛì¸"
            gSL_Select = gSL_Select & " FROM ïiî‘éÊà¯êÊÉ}ÉXÉ^ HT"
            gSL_Select = gSL_Select & " LEFT OUTER JOIN ïiî‘É}ÉXÉ^ HN"
            gSL_Select = gSL_Select & " ON   HN.ïiî‘ = HT.ïiî‘"
            gSL_Select = gSL_Select & " WHERE HT.ïiñ⁄ = 1"
            gSL_Select = gSL_Select & " AND   HT.ÉAÉhÉåÉX = '" & iB3_Adress & "'"
            gSL_Select = gSL_Select & " AND"
            gSL_Select = gSL_Select & " NOT"
            gSL_Select = gSL_Select & "     ( HT.ïiî‘ = '" & RTrim$(iH1_Hinbn) & "'"
            gSL_Select = gSL_Select & " AND   HT.éÊà¯êÊ = '" & RTrim$(iB2_Torcd) & "'"
            gSL_Select = gSL_Select & " AND   HT.éÛì¸ = '" & RTrim$(iB2_Ukeir) & "'"
            gSL_Select = gSL_Select & " )"
            If HTMRead(gSL_Select, 1) Then
                With CisFun
                    .MB_Lines = 6
                    .MB_MSG(2) = "    ìØàÍÇÃÉAÉhÉåÉXÇ™ë∂ç›ÇµÇ‹Ç∑ÅB               "
                    .MB_MSG(4) = "    ïi î‘  ÅF" & HTM.ï\é¶ïiî‘
                    .MB_MSG(5) = "    éÊà¯êÊÅF" & HTM.éÊà¯êÊ
                    .MB_MSG(6) = "    éÛ ì¸  ÅF" & HTM.éÛì¸
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                    iB3_Adress.SetFocus
                    Exit Function
                End With
            End If
'                       << çﬁóø >>
'
'            gSL_Select = "SELECT ZR.ï\é¶ê°ñ@, ZT.éÊà¯êÊ, ZT.éÛì¸, ZT.çﬁéø"
'            gSL_Select = gSL_Select & " FROM çﬁóøéÊà¯êÊÉ}ÉXÉ^ ZT"
'            gSL_Select = gSL_Select & " LEFT OUTER JOIN çﬁóøÉ}ÉXÉ^ ZR"
'            gSL_Select = gSL_Select & " ON  ZR.çﬁóøä«óùî‘çÜ = ZT.çﬁóøä«óùî‘çÜ"
'            gSL_Select = gSL_Select & " WHERE ÉAÉhÉåÉX = '" & iB3_Adress & "'"
'            If ZTMRead(gSL_Select, 1) Then
'                With CisFun
'                    .MB_Lines = 8
'                    .MB_MSG(2) = "    ìØàÍÇÃÉAÉhÉåÉXÇ™ë∂ç›ÇµÇ‹Ç∑ÅB               "
'                    .MB_MSG(4) = "    çﬁ éø ÅF" & ZTM.çﬁéø
'                    .MB_MSG(5) = "    ê° ñ@ ÅF" & ZTM.ï\é¶ê°ñ@
'                    .MB_MSG(6) = "    éËîzêÊ ÅF" & ZTM.éÊà¯êÊ
'                    .MB_MSG(7) = "    éÛ ì¸  ÅF" & ZTM.éÛì¸
'                    .MB_Title = ""
'                    .MB_Button = Error
'                    .MBOX
'                    iB3_Adress.SetFocus
'                    Exit Function
'                End With
'            End If
        Else
            iB3_Adress = ""
        End If
'
'-----     -----     -----     -----     -----     -----     -----     -----     < 2006/02/16 ADD  END  >
       B3lb_DenSyu = "": SaveKeitai = ""
       If Trim(iB3_DenSyu) <> "" Then
             If MeisyoGet("ì`ï[éÌóﬁî≠íç", iB3_DenSyu) Then
                 B3lb_DenSyu = SYM_Meisyo
                 SaveKeitai = CisFun.Left2(SYM.âpêîéö1, 1)
             Else
                With CisFun
                     .MB_Lines = 4
                     .MB_MSG(2) = "    ì`ï[éÌóﬁñ¢ìoò^            "
                     .MB_Title = ""
                     .MB_Button = Error
                     .MBOX
                End With
                iB3_DenSyu.SetFocus
                Exit Function
             End If
'''2007/12/11DEL             CisComboHNou.CodeGet
'''2007/12/11DEL             CisComboHJyu.CodeGet
'''2007/12/11DEL             If RTrim(iB3_DenSyu) = "2" And (CisComboHNou.Code <> "2" Or CisComboHJyu.Code <> "2") Then
'''2007/12/11DEL                With CisFun
'''2007/12/11DEL                    .MB_Lines = 4
'''2007/12/11DEL                    .MB_MSG(1) = "    ì`ï[éÌóﬁÇ™ ÇiÇ`ÇlÇ`ÇÃéûÅAî[ïièëå`ë‘ÅEéÛóÃèëå`ë‘ÇÕÉtÉ@ÉCÉãèoóÕå¿íËÇ≈Ç∑ÅB            "
'''2007/12/11DEL                    .MB_MSG(3) = "    î[ïièëå`ë‘ÅEéÛóÃèëå`ë‘ÇÉtÉ@ÉCÉãèoóÕÇ…ê›íËÇµÇ‹Ç∑ÅB            "
'''2007/12/11DEL                    .MB_Title = ""
'''2007/12/11DEL                    .MB_Button = Error
'''2007/12/11DEL                    .MBOX
'''2007/12/11DEL                End With
'''2007/12/11DEL                CisComboHNou.Code = "2"
'''2007/12/11DEL                CisComboHNou.TextGet
'''2007/12/11DEL                CisComboHJyu.Code = "2"
'''2007/12/11DEL                CisComboHJyu.TextGet
'''2007/12/11DEL            End If
'''2007/12/11DEL    '----<< É`ÉPÉbÉgí«â¡  >>-------------------------------------------------------------- 2005.01.31 (START)-------'
'''2007/12/11DEL            If RTrim(iB3_DenSyu) = "3" And CisComboHJyu.Code <> "" Then
'''2007/12/11DEL                With CisFun
'''2007/12/11DEL                    .MB_Lines = 4
'''2007/12/11DEL                    .MB_MSG(1) = "    ì`ï[éÌóﬁÇ™ é©ããÉ`ÉPÉbÉgÇÃéûÅAéÛóÃèëå`ë‘ÇÕèoóÕñ≥Çµå¿íËÇ≈Ç∑ÅB            "
'''2007/12/11DEL                    .MB_MSG(3) = "    éÛóÃèëå`ë‘ÇèoóÕñ≥ÇµÇ…ê›íËÇµÇ‹Ç∑ÅB            "
'''2007/12/11DEL                    .MB_Title = ""
'''2007/12/11DEL                    .MB_Button = Error
'''2007/12/11DEL                    .MBOX
'''2007/12/11DEL                End With
'''2007/12/11DEL                CisComboHJyu.Code = ""
'''2007/12/11DEL                CisComboHJyu.TextGet
'''2007/12/11DEL            End If
'''2007/12/11DEL            If RTrim(iB3_DenSyu) = "4" And (CisComboHNou.Code <> "2" Or CisComboHJyu.Code <> "") Then
'''2007/12/11DEL                With CisFun
'''2007/12/11DEL                    .MB_Lines = 4
'''2007/12/11DEL                    .MB_MSG(1) = "    ì`ï[éÌóﬁÇ™ éxããÉ`ÉPÉbÉgÇÃéûÅAî[ïièëå`ë‘ÇÕÉtÉ@ÉCÉãèoóÕÅEéÛóÃèëå`ë‘ÇÕèoóÕñ≥Çµå¿íËÇ≈Ç∑ÅB            "
'''2007/12/11DEL                    .MB_MSG(3) = "    î[ïièëå`ë‘ÇÉtÉ@ÉCÉãèoóÕÅEéÛóÃèëå`ë‘ÇèoóÕñ≥ÇµÇ…ê›íËÇµÇ‹Ç∑ÅB            "
'''2007/12/11DEL                    .MB_Title = ""
'''2007/12/11DEL                    .MB_Button = Error
'''2007/12/11DEL                    .MBOX
'''2007/12/11DEL                End With
'''2007/12/11DEL                CisComboHNou.Code = "2"
'''2007/12/11DEL                CisComboHNou.TextGet
'''2007/12/11DEL                CisComboHJyu.Code = ""
'''2007/12/11DEL                CisComboHJyu.TextGet
'''2007/12/11DEL            End If
'''2007/12/11DEL    '----<< É`ÉPÉbÉgí«â¡  >>-------------------------------------------------------------- 2005.01.31 (END  )-------'
'''2007/12/11DEL            If B3Comb_Nouhin.ListIndex = 0 And B3Comb_Jyuryo.ListIndex = 0 Then
'''2007/12/11DEL                With CisFun
'''2007/12/11DEL                    .MB_Lines = 4
'''2007/12/11DEL                    .MB_MSG(2) = "    î[ïièëå`ë‘ÅEéÛóÃèëå`ë‘ÇÇê›íËÇµÇƒâ∫Ç≥Ç¢ÅB            "
'''2007/12/11DEL                    .MB_Title = ""
'''2007/12/11DEL                    .MB_Button = Error
'''2007/12/11DEL                    .MBOX
'''2007/12/11DEL                End With
'''2007/12/11DEL               B3Comb_Nouhin.SetFocus
'''2007/12/11DEL               Exit Function
'''2007/12/11DEL            End If
            If (RTrim(iB3_DenSyu) = "3" Or RTrim(iB3_DenSyu) = "4") And (B3Chk_PrtkbS.Value = 0 Or B3Chk_PrtkbJ.Value = 0) Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(1) = "    ì`ï[éÌóﬁÇ™ É`ÉPÉbÉgÇÃéûÅAéxããèëî≠çsÅEéÛóÃèëî≠çsÇÕî≠çsñ≥Çµå¿íËÇ≈Ç∑ÅB            "
                    .MB_MSG(3) = "    éxããèëî≠çsÅEéÛóÃèëî≠çsÇî≠çsñ≥ÇµÇ…ê›íËÇµÇ‹Ç∑ÅB            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                B3Chk_PrtkbS.Value = 1
                B3Chk_PrtkbJ.Value = 1
            End If
        Else
'''2007/12/11DEL            B3Comb_Nouhin.ListIndex = 0
'''2007/12/11DEL            B3Comb_Jyuryo.ListIndex = 0
            B3Chk_PrtkbN.Value = 1
            B3Chk_PrtkbS.Value = 1
            B3Chk_PrtkbJ.Value = 1
            B3Chk_Juryo.Value = 0
        End If
'-----     -----     -----     -----     -----     -----     -----     -----     < 2006/02/16 MNT START >
        If CisFun.Val2(iB3_PNo) > maxPNo Then
           With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    PNO(É|ÉPÉbÉgNO)ÇÕ" & maxPNo & "à»â∫Ç…Çƒê›íËÇµÇƒâ∫Ç≥Ç¢            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
           End With
           iB3_PNo.SetFocus
           Exit Function
        End If
'-----     -----     -----     -----     -----     -----     -----     -----     < 2006/02/16 MNT  END  >
    End If
'========================================== 2009/07/22 Start
' çÏã∆É`ÉFÉbÉN
'    With vsGridS 'CisVsGridS
'        ' ãÛîíçsçÌèú
'        For mIx = 1 To .Rows - 1
'            If Trim(.TextMatrix(mIx, 0)) = "" And _
'               Trim(.TextMatrix(mIx, 1)) = "" And _
'               Trim(.TextMatrix(mIx, 3)) = "" Then
'                For mIx2 = mIx + 1 To .Rows - 1
'                    If Trim(.TextMatrix(mIx2, 0)) <> "" Or _
'                       Trim(.TextMatrix(mIx2, 1)) <> "" Or _
'                       Trim(.TextMatrix(mIx2, 3)) <> "" Then
'                       .TextMatrix(mIx, 0) = .TextMatrix(mIx2, 0)
'                       .TextMatrix(mIx, 1) = .TextMatrix(mIx2, 1)
'                       .TextMatrix(mIx, 2) = .TextMatrix(mIx2, 2)
'                       .TextMatrix(mIx, 3) = .TextMatrix(mIx2, 3)
'                       .TextMatrix(mIx2, 0) = ""
'                       .TextMatrix(mIx2, 1) = ""
'                       .TextMatrix(mIx2, 2) = ""
'                       .TextMatrix(mIx2, 3) = ""
'                       Exit For
'                    End If
'                Next mIx2
'            End If
'        Next mIx
'        ' ï™óﬁ(No)
'        For mIx = 1 To .Rows - 1
'            If Trim(.TextMatrix(mIx, 0)) = "" And _
'               Trim(.TextMatrix(mIx, 1)) = "" And _
'               Trim(.TextMatrix(mIx, 3)) = "" Then Exit For
'            For mIx2 = 1 To .Rows - 1
'                If Trim(.TextMatrix(mIx2, 0)) <> "" Or _
'                   Trim(.TextMatrix(mIx2, 1)) <> "" Or _
'                   Trim(.TextMatrix(mIx2, 3)) <> "" Then
'                   If CisFun.Val2(.TextMatrix(mIx, 0)) < CisFun.Val2(.TextMatrix(mIx2, 0)) Then
'                      For mIx3 = 0 To 3
'                          gStr = .TextMatrix(mIx2, mIx3)
'                          .TextMatrix(mIx2, mIx3) = .TextMatrix(mIx, mIx3)
'                          .TextMatrix(mIx, mIx3) = gStr
'                      Next mIx3
'                   End If
'                End If
'            Next mIx2
'        Next mIx
'        Set gObj = Nothing
'        For mIx = 1 To .Rows - 1
'            If Trim(.TextMatrix(mIx, 0)) = "" And Trim(.TextMatrix(mIx, 1)) = "" Then
'               .TextMatrix(mIx, 2) = ""
'               .TextMatrix(mIx, 3) = ""
'            Else
'               If Trim(.TextMatrix(mIx, 0)) = "" Then
'                  With CisFun
'                       .MB_Lines = 4
'                       .MB_MSG(2) = "    No(çÏã∆èá)Çì¸óÕÇµÇƒâ∫Ç≥Ç¢           "
'                       .MB_Title = ""
'                       .MB_Button = Error
'                       .MBOX
'                    End With
'                    Set gObj = iB3_No
'                    .Row = mIx
'                    GoTo Sagyo_Err
'               End If
'               If Trim(.TextMatrix(mIx, 1)) = "" Then
'                  With CisFun
'                       .MB_Lines = 4
'                       .MB_MSG(2) = "    çÏã∆Çì¸óÕÇµÇƒâ∫Ç≥Ç¢           "
'                       .MB_Title = ""
'                       .MB_Button = Error
'                       .MBOX
'                    End With
'                    Set gObj = iB3_Sagyo
'                    .Row = mIx
'                    GoTo Sagyo_Err
'               End If
'               .TextMatrix(mIx, 2) = ""
'               If MeisyoGet("çÏã∆ä«óù", .TextMatrix(mIx, 1)) Then
'                   .TextMatrix(mIx, 2) = SYM_Meisyo
'               Else
'                    With CisFun
'                       .MB_Lines = 4
'                       .MB_MSG(2) = "    çÏã∆ñ¢ìoò^            "
'                       .MB_Title = ""
'                       .MB_Button = Error
'                       .MBOX
'                    End With
'                    Set gObj = iB3_Sagyo
'                    .Row = mIx
'                    GoTo Sagyo_Err
'               End If
'            End If
'        Next mIx
''        ' Noèdï°É`ÉFÉbÉN
'        For mIx = 1 To .Rows - 1
'            If Trim(.TextMatrix(mIx, 0)) <> "" Then
'                For mIx2 = mIx + 1 To .Rows - 1
'                    If Trim(.TextMatrix(mIx2, 0)) <> "" Then
'                       If CisFun.Val2(.TextMatrix(mIx, 0)) = CisFun.Val2(.TextMatrix(mIx2, 0)) Then
'                          With CisFun
'                              .MB_Lines = 4
'                              .MB_MSG(2) = "    No(çÏã∆èáÅjÇ™èdï°ÇµÇƒÇ‹Ç∑ÅB            "
'                              .MB_Title = ""
'                              .MB_Button = Error
'                              .MBOX
'                           End With
'                           Set gObj = iB3_No
'                           .Row = mIx2
'                           GoTo Sagyo_Err
'                       End If
'                    End If
'                Next mIx2
'            End If
'        Next mIx
'    End With
''+--------------------------------------------------------------- 2004/12/13 Insert
'' ã@äBÉ`ÉFÉbÉN
'    gCnt1 = 0: gCnt2 = 0
'    With VsGridK
'        For mIx = 1 To .Rows - 1
'            If Trim(.TextMatrix(mIx, 1)) = "" Then
'               .TextMatrix(mIx, 2) = ""
'               .TextMatrix(mIx, 3) = ""
'            Else
'               If Trim(.TextMatrix(mIx, 1)) = "" Then
'                  With CisFun
'                       .MB_Lines = 4
'                       .MB_MSG(2) = "    ã@äBÇì¸óÕÇµÇƒâ∫Ç≥Ç¢           "
'                       .MB_Title = ""
'                       .MB_Button = Error
'                       .MBOX
'                    End With
'                    .Row = mIx
'                    GoTo Kikai_Err
'               End If
'               .TextMatrix(mIx, 2) = ""
'               If KMMMeisyoGet(.TextMatrix(mIx, 1)) Then
'                   .TextMatrix(mIx, 2) = KMM_Meisyo
'                   If Trim(.TextMatrix(mIx, 3)) <> "" Then
'                      gCnt1 = gCnt1 + 1
'                   End If
'                   If gCnt2 = 0 Then
'                      gCnt2 = mIx
'                   End If
'               Else
'                    With CisFun
'                       .MB_Lines = 4
'                       .MB_MSG(2) = "    ã@äBáÇñ¢ìoò^            "
'                       .MB_Title = ""
'                       .MB_Button = Error
'                       .MBOX
'                    End With
'                    .Row = mIx
'                    GoTo Kikai_Err
'               End If
'            End If
'        Next mIx
''        ' ã@äBáÇèdï°É`ÉFÉbÉN
'        For mIx = 1 To .Rows - 1
'            If Trim(.TextMatrix(mIx, 1)) <> "" Then
'                For mIx2 = mIx + 1 To .Rows - 1
'                    If Trim(.TextMatrix(mIx2, 1)) <> "" Then
'                       If Trim(.TextMatrix(mIx, 1)) = Trim(.TextMatrix(mIx2, 1)) Then
'                          With CisFun
'                              .MB_Lines = 4
'                              .MB_MSG(2) = "    ã@äBáÇÇ™èdï°ÇµÇƒÇ‹Ç∑ÅB            "
'                              .MB_Title = ""
'                              .MB_Button = Error
'                              .MBOX
'                           End With
'                           .Row = mIx2
'                           GoTo Kikai_Err
'                       End If
'                    End If
'                Next mIx2
'            End If
'        Next mIx
'        If gCnt1 = 0 And gCnt2 <> 0 Then
'           .TextMatrix(gCnt2, 3) = "Åõ"
'        End If
'    End With
'+--------------------------------------------------------------- 2004/12/13 Insert End
'========================================== 2009/07/22 End
    
    
    Body3Chk = True
    Exit Function
Sagyo_Err:
'====================================== 2009/07/22 Start
'    Call CisVsGridS.EditEnable
'    Call CisVsGridS.EditVisible
'    gObj.SetFocus
'====================================== 2009/07/22 End
    Exit Function
'+--------------------------------------------------------------- 2004/12/13 Insert
Kikai_Err:
'====================================== 2009/07/22 Start
'    Call CisVsGridK.EditEnable
'    Call CisVsGridK.EditVisible
'    iB3_Kikai.SetFocus
'====================================== 2009/07/22 End
    Exit Function
'+--------------------------------------------------------------- 2004/12/13 Insert End
End Function

'**************************************
'*                                    *
'*       ÉfÅ[É^ÉxÅ[ÉXì¸óÕ Åï ï\é¶       *
'*                                    *
'**************************************
'+---------------------------------+
'+      äÓñ{èÓïÒ
'+---------------------------------+
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    Call HeadBodyClear("B")
    ' ç›å…î≠íçÅ@ÉRÉìÉ{ÉZÉbÉg
    Call CisComboZaiko.ComboSet

    If Not HNMRead(gSL_Select, 1) Then
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
        iH1_Hinbn = RTrim(HNM.ïiî‘)
        Call BodySet("")
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
Private Sub BodySet(SetType As String)
    
    With HNM
        iB1_Syakei = .é‘å^
        iB1_SSeban = .é–ì‡îwî‘
        iB1_HinNM = .ïiñº
        iB1_KSyuyo = .äÓñ{é˚óeêî
        iB1_KUtsuwa = .äÓñ{é˚óeäÌ
        '''iB1_Kotei = .çHíˆ
        '''Call iB1_Kotei_LostFocus
        '''iB1_NKotei = .éüçHíˆ
        '''Call iB1_NKotei_LostFocus
        iB1_Tanto = .íSìñ
        Call iB1_Tanto_LostFocus
        iB1_Edit = .ïiî‘ï“èW
        iB1_KataNo = .å^NO
        iB1_KataTor = .å^èäç›ín
        Call iB1_KataTor_LostFocus
        iB1_Image = .IMAGE
        ' ç›å…ä«óù
        With CisComboZaiko
            .Code = HNM.ç›å…ä«óùãÊï™
            .TextGet
        End With
        Call B1Comb_Zaiko_Click
        ' ç›å…åWêî
        iB1_Zaiko = CisFun.RSetFld(.ç›å…åWêî, 0, iB1_Zaiko.cFormat)
        ' ç›å…ì˙êî
        iB1_ZaikoDay = CisFun.RSetFld(.ç›å…ì˙êî, 0, iB1_ZaikoDay.cFormat)
        ' ç›å…î≠íçì_
        iB1_ZaikoL = CisFun.RSetFld(.ç›å…î≠íçì_, 0, iB1_ZaikoL.cFormat)
        ' é¿êîãÊï™
        B1Chk_JSKbn.Value = .é¿êîãÊï™
        Call B1Chk_JSKbn_Click
        
        iB1_Hinku = .ïiî‘ãÊï™
        Call iB1_Hinku_LostFocus
        'ABCãÊï™
        iB1_Abc = .ABC      ' 2004/04/15 Mnt
        
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
        B1Op_Buhin(.ïîïiãÊï™).Value = True
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
'-----------------------------------------------------------' 2009.10.20 add start
        iB1_DaiCD = .ìæà”êÊ             ' ë„ï\ìæà”êÊ
        Call iB1_DaiCD_LostFocus
'-----------------------------------------------------------' 2009.10.20 add end

    End With
    ' ïiî‘ï“èWï\é¶
    B1lb_EHinbn = EditHinbn(iH1_Hinbn, iB1_Edit)
'    ' çﬁóøèÓïÒ
    Call ZairyoKensu

    ' éÊà¯êÊèÓïÒ
    gSL_Select = "Select * from ïiî‘éÊà¯êÊÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " Where ïiî‘ = '" & RTrim(HNM.ïiî‘) & "'"
    gSL_Select = gSL_Select & " Order by éÊà¯êÊ,éÛì¸,ïiñ⁄"
    If HTMRead(gSL_Select, , 1) Then
       Call GridSet
       iB2_Torcd.Tag = 1
       Call CisComboTehai.ComboSet
'''2007/12/11DEL       Call CisComboHNou.ComboSet
'''2007/12/11DEL       Call CisComboHJyu.ComboSet
       Call CisComboHMei.ComboSet
'''2007/12/11DEL       Call CisComboJNou.ComboSet
'''2007/12/11DEL       Call CisComboJJyu.ComboSet
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
       Call CisComboJsya.ComboSet
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
       
       Call GridToText
    
'========================================================== *** 2009/07/22 Start
       If ProcHB = "B3" Then
          Call Kojyun_SQL_Set
          If HKTRead(gSL_Select, , 1) Then
''''             Call GridSet3              ' 2009.10.16 del
            'SAVE_Select = gSL_Select       ' 2009.10.16 add
          End If
          
          If HKTRead(gSL_Select, , 1) Then
             Call GridSet2
          End If
       End If
'========================================================== *** 2009/07/22 End

    End If
    Call HTMClose

End Sub
'========================================================== *** 2009/07/22 Start
Private Sub Kojyun_SQL_Set()
    ' çHèáèÓïÒ
    gSL_Select = "SELECT HK.*,ISNULL(SY1.ílñºèÃ,'') ã@äBñº,ISNULL(SY2.ílñºèÃ,'') ã@äBNOñº,"
    gSL_Select = gSL_Select & " ISNULL(SY3.ílñºèÃ,'') çHíˆñº,ISNULL(TR.ó™èÃ,'') äOíçâ¡çHñº "
    gSL_Select = gSL_Select & " From ïiî‘çHèáÉ}ÉXÉ^ HK"
    gSL_Select = gSL_Select & " Left Outer join ñºèÃÉ}ÉXÉ^ SY1"
    gSL_Select = gSL_Select & " On  SY1.ãÊï™ñºèÃ = 'ã@äB' "
    gSL_Select = gSL_Select & " And SY1.íl = HK.ã@äB"
    gSL_Select = gSL_Select & " And SY1.ãÊï™É^ÉCÉv = ''"
    gSL_Select = gSL_Select & " Left Outer join ñºèÃÉ}ÉXÉ^ SY2"
    gSL_Select = gSL_Select & " On  SY2.ãÊï™ñºèÃ = 'ã@äBNO' "
    gSL_Select = gSL_Select & " And SY2.íl = HK.ã@äBNO"
    gSL_Select = gSL_Select & " And SY2.ãÊï™É^ÉCÉv = ''"
    gSL_Select = gSL_Select & " Left Outer join ñºèÃÉ}ÉXÉ^ SY3"
    gSL_Select = gSL_Select & " On  SY3.ãÊï™ñºèÃ = 'çHíˆä«óù' "
    gSL_Select = gSL_Select & " And SY3.íl = HK.çHíˆ"
    gSL_Select = gSL_Select & " And SY3.ãÊï™É^ÉCÉv = ''"
    gSL_Select = gSL_Select & " Left Outer join éÊà¯êÊÉ}ÉXÉ^ TR"
    gSL_Select = gSL_Select & " On  TR.éÊà¯êÊCD = HK.äOíçâ¡çHêÊCD "
    gSL_Select = gSL_Select & " And TR.éÊà¯êÊãÊï™ = 1"
        
    gSL_Select = gSL_Select & " WHERE HK.ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   HK.éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
    gSL_Select = gSL_Select & " And   HK.éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
    If B2Op_Hinmoku(0) Then
       gSL_Select = gSL_Select & " And   HK.ïiñ⁄ = 0"
    Else
       gSL_Select = gSL_Select & " And   HK.ïiñ⁄ = 1"
    End If
    gSL_Select = gSL_Select & " ORDER BY HK.çHèá"
End Sub
Private Sub Kojyun_SQL_Set2()
    ' çHèáèÓïÒ(vsGrid2óp)
    gSL_Select = "SELECT HK.*,ISNULL(SY1.ílñºèÃ,'') ã@äBñº,ISNULL(SY2.ílñºèÃ,'') ã@äBNOñº,"
    gSL_Select = gSL_Select & " ISNULL(SY3.ílñºèÃ,'') çHíˆñº,ISNULL(TR.ó™èÃ,'') äOíçâ¡çHñº "
    gSL_Select = gSL_Select & " From ïiî‘çHèáÉ}ÉXÉ^ HK"
    gSL_Select = gSL_Select & " Left Outer join ñºèÃÉ}ÉXÉ^ SY1"
    gSL_Select = gSL_Select & " On  SY1.ãÊï™ñºèÃ = 'ã@äB' "
    gSL_Select = gSL_Select & " And SY1.íl = HK.ã@äB"
    gSL_Select = gSL_Select & " And SY1.ãÊï™É^ÉCÉv = ''"
    gSL_Select = gSL_Select & " Left Outer join ñºèÃÉ}ÉXÉ^ SY2"
    gSL_Select = gSL_Select & " On  SY2.ãÊï™ñºèÃ = 'ã@äBNO' "
    gSL_Select = gSL_Select & " And SY2.íl = HK.ã@äBNO"
    gSL_Select = gSL_Select & " And SY2.ãÊï™É^ÉCÉv = ''"
    gSL_Select = gSL_Select & " Left Outer join ñºèÃÉ}ÉXÉ^ SY3"
    gSL_Select = gSL_Select & " On  SY3.ãÊï™ñºèÃ = 'çHíˆä«óù' "
    gSL_Select = gSL_Select & " And SY3.íl = HK.çHíˆ"
    gSL_Select = gSL_Select & " And SY3.ãÊï™É^ÉCÉv = ''"
    gSL_Select = gSL_Select & " Left Outer join éÊà¯êÊÉ}ÉXÉ^ TR"
    gSL_Select = gSL_Select & " On  TR.éÊà¯êÊCD = HK.äOíçâ¡çHêÊCD "
    gSL_Select = gSL_Select & " And TR.éÊà¯êÊãÊï™ = 1"
        
    gSL_Select = gSL_Select & " WHERE HK.ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   HK.éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
    gSL_Select = gSL_Select & " And   HK.éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
    If B2Op_Hinmoku(0) Then
       gSL_Select = gSL_Select & " And   HK.ïiñ⁄ = 0"
    Else
       gSL_Select = gSL_Select & " And   HK.ïiñ⁄ = 1"
    End If
    gSL_Select = gSL_Select & " ORDER BY HK.çHèá"
End Sub
'========================================================== *** 2009/07/22 End
'**************************************
'*                                    *
'*          ÉfÅ[É^ÉxÅ[ÉXçXêV           *
'*                                    *
'**************************************
'+----------------------------+
'+      çÌ èú èà óù (ïiî‘)     +
'+----------------------------+
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
        ' ïiî‘É}ÉXÉ^
        .SQL = "DELETE FROM ïiî‘É}ÉXÉ^ "
        .SQL = .SQL & " WHERE ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        .DBExec
        ' ïiî‘éÊà¯êÊÉ}ÉXÉ^
        .SQL = "DELETE FROM ïiî‘éÊà¯êÊÉ}ÉXÉ^ "
        .SQL = .SQL & " WHERE ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        On Error GoTo DBDelete_Err
        .DBExec
        ' ïiî‘éÊà¯çÏã∆É}ÉXÉ^
        .SQL = "DELETE FROM ïiî‘éÊà¯çÏã∆É}ÉXÉ^ "
        .SQL = .SQL & " WHERE ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        On Error GoTo DBDelete_Err
        .DBExec
    '------------------------------------------------- 2006/03/28 Insert
        ' ïiî‘éÊà¯ã@äBÉ}ÉXÉ^
        .SQL = "DELETE FROM ïiî‘éÊà¯ã@äBÉ}ÉXÉ^ "
        .SQL = .SQL & " Where ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        .DBExec
        ' Ç©ÇÒÇŒÇÒó¨ìÆä«óùÉeÅ[ÉuÉã
        .SQL = "DELETE FROM Ç©ÇÒÇŒÇÒó¨ìÆä«óùÉeÅ[ÉuÉã "
        .SQL = .SQL & " WHERE ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        .DBExec
    '------------------------------------------------- 2006/03/28 Insert End
        
    '================================================= *** 2009/07/22 STart
        ' ïiî‘çHèáÉ}ÉXÉ^
        .SQL = "DELETE FROM ïiî‘çHèáÉ}ÉXÉ^ "
        .SQL = .SQL & " WHERE ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        On Error GoTo DBDelete_Err
        .DBExec
        ' ïiî‘çHèáÉZÉbÉgïiî‘É}ÉXÉ^
        .SQL = "DELETE FROM ïiî‘çHèáÉZÉbÉgïiî‘É}ÉXÉ^ "
        .SQL = .SQL & " WHERE ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        On Error GoTo DBDelete_Err
        .DBExec
    '================================================= *** 2009/07/22 STart
        
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
'+----------------------------+
'+      çÌ èú èà óù (éÊà¯êÊ)   +
'+----------------------------+
Private Function DBDeleteTor()
    DBDeleteTor = False
'  ( ämîFÉÅÉbÉZÅ[ÉW )
    K_Sykbnm2.Caption = SyoriNM2
    K_Sykbnm2.Visible = True
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     éÊà¯êÊèÓïÒÇÃçÌèúèàóùÇçsÇ¢Ç‹Ç∑      "
        .MB_Title = "çÌèúèàóù"
        .MB_MSG(3) = "     ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    With CisDB
        On Error GoTo DBDeleteTor_Err
        .DBTran (TransBegin) ' ƒ◊›ªﬁ∏ºÆ›äJén
' ( çÌèúèàóù )
        ' ïiî‘éÊà¯êÊÉ}ÉXÉ^
        .SQL = "DELETE FROM ïiî‘éÊà¯êÊÉ}ÉXÉ^ "
        .SQL = .SQL & " Where ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        .SQL = .SQL & " And   éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
        .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
        On Error GoTo DBDeleteTor_Err
        .DBExec
        ' ïiî‘éÊà¯çÏã∆É}ÉXÉ^
        .SQL = "DELETE FROM ïiî‘éÊà¯çÏã∆É}ÉXÉ^ "
        .SQL = .SQL & " Where ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        .SQL = .SQL & " And   éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
        .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
        On Error GoTo DBDeleteTor_Err
        .DBExec
    '------------------------------------------------- 2006/03/28 Insert
        ' ïiî‘éÊà¯ã@äBÉ}ÉXÉ^
        .SQL = "DELETE FROM ïiî‘éÊà¯ã@äBÉ}ÉXÉ^ "
        .SQL = .SQL & " Where ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        .SQL = .SQL & " And   éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
        .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
        .DBExec
        ' Ç©ÇÒÇŒÇÒó¨ìÆä«óùÉeÅ[ÉuÉã
        .SQL = "DELETE FROM Ç©ÇÒÇŒÇÒó¨ìÆä«óùÉeÅ[ÉuÉã "
        .SQL = .SQL & " WHERE ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        .SQL = .SQL & " And   éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
        .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
        .DBExec
    '------------------------------------------------- 2006/03/28 Insert End
        
    '================================================= 2009/07/22 Start
        ' ïiî‘çHèáÉ}ÉXÉ^
        .SQL = "DELETE FROM ïiî‘çHèáÉ}ÉXÉ^ "
        .SQL = .SQL & " WHERE ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        .SQL = .SQL & " And   éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
        .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
        .DBExec
        ' ïiî‘çHèáÉZÉbÉgïiî‘É}ÉXÉ^
        .SQL = "DELETE FROM ïiî‘çHèáÉZÉbÉgïiî‘É}ÉXÉ^ "
        .SQL = .SQL & " WHERE ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        .SQL = .SQL & " And   éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
        .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
        .DBExec
    '================================================= 2009/07/22 End
    
        .DBTran (TransCommit) ' ƒ◊›ªﬁ∏ºÆ› èIóπ
        
        gLong = CisFun.Val2(iB2_Torcd.Tag)
        With vsGrid1
             .RemoveItem gLong
             If .Rows < CisVsGrid.DispMax Then
                .Rows = CisVsGrid.DispMax
             End If
        End With
        CisVsGrid.ReNumber
        On Error GoTo 0
    End With
    DBDeleteTor = True
    Exit Function
DBDeleteTor_Err:
    CisFun.ErrorBox
    End
End Function

'+----------------------+
'+    í«â¡•çXêV èàóù     +
'+----------------------+
Private Function DBPut(Optional Msg As Boolean = True)
    DBPut = False
    Dim AddFlg  As Boolean

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    If Msg Then
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
            If Not .MBOX Then
               If Trim(vsGrid1.TextMatrix(vsGrid1.Row, 1)) <> "" Then
                  vsGrid1.SetFocus
               End If
               Exit Function
            End If
        End With
    End If
    On Error GoTo DBPut_Err
'   É}ÉXÉ^çÄñ⁄à⁄ëó
    
    gSL_Select = "SELECT * From ïiî‘É}ÉXÉ^"
    gSL_Select = gSL_Select & " WHERE ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    If HNMRead(gSL_Select, 1) Then
       AddFlg = False
    Else
       Call ItemsClearHNM
       AddFlg = True
       HNM.ïiî‘ = RTrim(iH1_Hinbn)
    End If
    With HNM
        .é‘å^ = RTrim(iB1_Syakei)
        .é–ì‡îwî‘ = RTrim(iB1_SSeban)
        .ïiñº = RTrim(iB1_HinNM)
        .äÓñ{é˚óeêî = CisFun.Val2(iB1_KSyuyo)
        .äÓñ{é˚óeäÌ = RTrim(iB1_KUtsuwa)
        '''.çHíˆ = iB1_Kotei
        '''.éüçHíˆ = iB1_NKotei
        .íSìñ = iB1_Tanto
        .ïiî‘ï“èW = iB1_Edit
        .ï\é¶ïiî‘ = RTrim(B1lb_EHinbn)
        .IMAGE = RTrim(iB1_Image)
        .å^NO = RTrim(iB1_KataNo)
        .å^èäç›ín = RTrim(iB1_KataTor)
        .ABC = RTrim(iB1_Abc)
        .ïiî‘ãÊï™ = RTrim(iB1_Hinku)
        .é¿êîãÊï™ = B1Chk_JSKbn.Value
        .âºãÊï™ = 0
        With CisComboZaiko
            .CodeGet
            HNM.ç›å…ä«óùãÊï™ = CisFun.Val2(.Code)
        End With
        .ç›å…åWêî = CisFun.Val2(iB1_Zaiko)
        .ç›å…î≠íçì_ = CisFun.Val2(iB1_ZaikoL)
        .ç›å…ì˙êî = CisFun.Val2(iB1_ZaikoDay)
        Select Case .ç›å…ä«óùãÊï™
             Case 2: .ç›å…ä«óùêîíl = .ç›å…åWêî
             Case 3: .ç›å…ä«óùêîíl = .ç›å…î≠íçì_
             Case 4: .ç›å…ä«óùêîíl = .ç›å…ì˙êî
        End Select
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
        .ïîïiãÊï™ = B1Op_Buhin(0).Tag
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
'-----------------------------------------------------------' 2009.10.20 add start
        .ìæà”êÊ = RTrim(iB1_DaiCD)      ' ë„ï\ìæà”êÊ
'-----------------------------------------------------------' 2009.10.20 add end
    End With

    Call CisDB.DBTran(TransBegin)  ' ƒ◊›ªﬁ∏ºÆ›äJén
' ( í«â¡èàóù )
    If AddFlg Then
        HNM.çÏê¨é“ = gTanto
        Call HNMInsert
        iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
    Else
' ( çXêVèàóù )
        HNM.çXêVé“ = gTanto
        Call HNMUpdate
        iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
    End If

    Call DBPutYouki(iB1_KUtsuwa)
'
    CisDB.DBTran (TransCommit) ' ƒ◊›ªﬁ∏ºÆ› èIóπ
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'*************************************************
'*      éÊà¯êÊèÓïÒçXêV                            *
'*************************************************
Public Function DBPutTor() As Boolean
    Dim wAdress     As String
    
    DBPutTor = False
    
    Dim wRow    As Long
    Dim AddFlg  As Boolean
    
    K_Sykbnm2.Caption = SyoriNM2
    K_Sykbnm2.Visible = True
    
    With CisFun
        .MB_Lines = 4
        If SyoriKB2 = "ADD" Then
            .MB_MSG(2) = "     éÊà¯êÊèÓïÒÇÃí«â¡èàóùÇçsÇ¢Ç‹Ç∑      "
            .MB_Title = "í«â¡èàóù"
        Else
            .MB_MSG(2) = "     éÊà¯êÊèÓïÒÇÃèCê≥èàóùÇçsÇ¢Ç‹Ç∑      "
            .MB_Title = "èCê≥èàóù"
        End If
        .MB_MSG(3) = "     ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPutTor_Err
    ' ±ƒﬁ⁄Ωè»ó™éû,±ƒﬁ⁄Ωé©ìÆçÃî‘
    wAdress = ""
    If CisFun.Val2(iB3_Adress) = 0 And B3Chk_Auto.Value = 1 Then
       gSL_Select = "Select Max(ÉAÉhÉåÉX) ÉAÉhÉåÉX From ïiî‘éÊà¯êÊÉ}ÉXÉ^"
       Call HTMRead(gSL_Select, 1)
       If CisFun.Val2(HTM.ÉAÉhÉåÉX) = 0 Then
          wAdress = Format(1, iB3_Adress.cFormat)
       Else
          gStr = String(iB3_Adress.MaxLength, "9")
          gLong = CisFun.Val2(HTM.ÉAÉhÉåÉX) + 1
          gInt = 0
          Do Until gInt > 1
             gSL_Select = "Select ÉAÉhÉåÉX From ïiî‘éÊà¯êÊÉ}ÉXÉ^"
             gSL_Select = gSL_Select & " Where ÉAÉhÉåÉX = '" & Format(gLong, iB3_Adress.cFormat) & "'"
             If Not HTMRead(gSL_Select, 1) Then
                wAdress = Format(gLong, iB3_Adress.cFormat)
                Exit Do
             End If
             gLong = gLong + 1
             If gLong > CisFun.Val2(gStr) Then
                gLong = 1
                gInt = gInt + 1
             End If
          Loop
       End If
       
    End If
    With vsGrid1
        .Redraw = False
        If Trim(iB2_Torcd.Tag) = "" Then
            wRow = 0
            For mIx = 1 To .Rows - 1
                If Trim(.TextMatrix(mIx, 1)) = "" Then
                   wRow = mIx
                   Exit For
                End If
            Next mIx
        Else
            wRow = CisFun.Val2(iB2_Torcd.Tag)
            For mIx = 0 To .Cols - 1
                .TextMatrix(wRow, mIx) = ""
            Next mIx
        End If
        If wRow = 0 Then
           .Rows = .Rows + 1
           wRow = .Rows - 1
        End If
        .Row = wRow
        
        gSL_Select = "Select * from ïiî‘éÊà¯êÊÉ}ÉXÉ^ "
        gSL_Select = gSL_Select & " Where ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " And   éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
        gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
        If HTMRead(gSL_Select, 1) Then
           AddFlg = False
        Else
           Call ItemsClearHTM
           AddFlg = True
           HTM.ïiî‘ = RTrim(iH1_Hinbn)
           HTM.éÊà¯êÊ = RTrim(iB2_Torcd)
           HTM.éÛì¸ = RTrim(iB2_Ukeir)
        End If
        With vsGrid1
            If CisFun.Val2(iB3_Adress) = 0 And Trim(wAdress) <> "" Then
               HTM.ÉAÉhÉåÉX = wAdress
               iB3_Adress = wAdress
            Else
               If CisFun.Val2(iB3_Adress) <> 0 Then
                  HTM.ÉAÉhÉåÉX = Format(CisFun.Val2(iB3_Adress), iB3_Adress.cFormat)
               Else
                  HTM.ÉAÉhÉåÉX = ""
               End If
            End If
            If B2Op_Hinmoku(0).Value Then
               HTM.ïiñ⁄ = 0
            Else
               HTM.ïiñ⁄ = 1
            End If
            HTM.îwî‘çÜ = iB3_Seban
            If B3Op_SK(0).Value Then
               HTM.SKãÊï™ = 0
            Else
               HTM.SKãÊï™ = 1
            End If
            If CisFun.Val2(iB3_Cycle1) <> 0 Or CisFun.Val2(iB3_Cycle2) <> 0 Or _
               CisFun.Val2(iB3_Cycle3) <> 0 Then
               HTM.ÉTÉCÉNÉã = Format(iB3_Cycle1, "0") & _
                             Format(iB3_Cycle2, "00") & _
                             Format(iB3_Cycle3, "00")
'*--------------------------------------------------<< 2004/09/24 >> Insert Start
            Else
                HTM.ÉTÉCÉNÉã = ""
'*--------------------------------------------------<< 2004/09/24 >> Insert End
            End If
            HTM.é˚óeêî1 = CisFun.Val2(iB3_Syuyo1)
            HTM.é˚óeäÌ1 = iB3_Youki1
            HTM.é˚óeêî2 = CisFun.Val2(iB3_Syuyo2)
            HTM.é˚óeäÌ2 = iB3_Youki2
            HTM.ïœçXì˙ = iB3_HYmd
            If Trim(iB3_HBin) <> "" Then
               HTM.ïœçXï÷ = Format(iB3_HBin, "00")
            End If
            HTM.íuèÍ = iB3_Okiba
            HTM.å_ñÒêÊ = iB3_KTorcd
            ' î≠íç
            If B2Op_Hinmoku(1).Value Then
               HTM.Ç©åWêî = CisFun.Val2(iB3_KKeisu)
               HTM.î≠íçî‰ó¶ = CisFun.Val2(iB3_Hiri)
               HTM.óLå¯ä˙å¿äJén = iB3_YSYmd
               HTM.óLå¯ä˙å¿èIóπ = iB3_YEYmd
               With CisComboTehai
                   .CodeGet
                   HTM.éËîzãÊï™ = .Code
               End With
               
               HTM.î[ì¸êÊ = iB3_NTorcd
               HTM.î[ì¸êÊéÛì¸ = iB3_Ukeir
            ' ë„çs
                If HTM.éËîzãÊï™ = "3" Then
                    HTM.ë„çsãÊï™ = ChkB3_Daiko.Value
                    HTM.ë„çså_ñÒêÊ = iB3_DaikoK
                Else
                    HTM.ë„çsãÊï™ = 0
                    HTM.ë„çså_ñÒêÊ = ""
                End If
                                
               HTM.ïΩèÄâªãÊï™ = CisFun.Val2(B3Op_Heijun(0).Tag)
               HTM.ïΩèÄâªåWêî = CisFun.Val2(iB3_HKeisu)
               HTM.åWêî1 = CisFun.Val2(iB3_Keisu1)
               HTM.åWêî2 = CisFun.Val2(iB3_Keisu2)
               HTM.åWêî3 = CisFun.Val2(iB3_Keisu3)
               HTM.LT = CisFun.Val2(iB3_LT)
               HTM.ÉçÉbÉg1 = CisFun.Val2(iB3_LSu)
               HTM.ÉçÉbÉg2 = CisFun.Val2(iB3_LSu2)
               HTM.Week1 = B3Chk_Week(0).Value
               HTM.Week2 = B3Chk_Week(1).Value
               HTM.Week3 = B3Chk_Week(2).Value
               HTM.Week4 = B3Chk_Week(3).Value
               HTM.Week5 = B3Chk_Week(4).Value
               HTM.Day1 = CisFun.Val2(iB3_Day(0))
               HTM.Day2 = CisFun.Val2(iB3_Day(1))
               HTM.Day3 = CisFun.Val2(iB3_Day(2))
               HTM.Day4 = CisFun.Val2(iB3_Day(3))
               HTM.Day5 = CisFun.Val2(iB3_Day(4))
               HTM.ì‡é¶î≠íç = B3Chk_Naiji.Value
               HTM.äÓèÄçHêî = CisFun.Val2(iB3_Kosu)
               HTM.îıçl = iB3_Biko
               HTM.PNO = CisFun.Val2(iB3_PNo)
               HTM.ì`ï[éÌóﬁ = iB3_DenSyu
               
               HTM.î[ïièëå`ë‘ = RTrim(SaveKeitai)
               HTM.éÛóÃèëå`ë‘ = RTrim(SaveKeitai)
'''2007/12/11DEL               With CisComboHNou
'''2007/12/11DEL                   .CodeGet
'''2007/12/11DEL                   HTM.î[ïièëå`ë‘ = .Code
'''2007/12/11DEL               End With
'''2007/12/11DEL               With CisComboHJyu
'''2007/12/11DEL                   .CodeGet
'''2007/12/11DEL                   HTM.éÛóÃèëå`ë‘ = .Code
'''2007/12/11DEL               End With
               With CisComboHMei
                   .CodeGet
                   HTM.ñæç◊èëî≠çs = .Code
               End With
               '-----* 2008/06/09 Start
               'HTM.î[ïièëî≠çsãÊï™ = B3Chk_PrtkbN.Value
               If HTM.éËîzãÊï™ <> 4 Then
                  HTM.î[ïièëî≠çsãÊï™ = B3Chk_PrtkbN.Value
               Else
                  HTM.î[ïièëî≠çsãÊï™ = B3Chk_PrtkbSN.Value
               End If
               '-----* 2008/06/09 End
               HTM.éxããèëî≠çsãÊï™ = B3Chk_PrtkbS.Value
               HTM.éÛóÃèëî≠çsãÊï™ = B3Chk_PrtkbJ.Value
               HTM.éÛóÃèëî≠çs = B3Chk_Juryo.Value
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
               If B3Chk_Gouho.Value = 0 Then
                  HTM.çÜï‚ãÊï™ = "G"
               Else
                  HTM.çÜï‚ãÊï™ = "H"
               End If
               HTM.ëOçHíˆãÊï™ = B3Chk_MaeHak.Value
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
               '=============================================== 2009/01/19 Start
               If HTM.éËîzãÊï™ = 1 Then
                  HTM.íºëóì`ï[î≠çsãÊï™ = B3Chk_PrtkbSN1.Value
               End If
               '=============================================== 2009/01/19 End
               '====================== *** 2009/07/06 Start
               HTM.èoâ◊èÍ = ""
               '====================== *** 2009/07/06 End
            
               '====================== *** 2009/07/22 Start
               HTM.ïîïiãÊï™ = CisFun.Val2(B3Op_HinKu(0).Tag)
               HTM.äÓèÄíPâø = CisFun.Val2(iB3_Kijyun)
               HTM.éwé¶èëî≠çsãÊï™ = 0
               '====================== *** 2009/07/22 End
             
            Else
               HTM.Ç©åWêî = 0
               HTM.î≠íçî‰ó¶ = 0
               HTM.óLå¯ä˙å¿äJén = ""
               HTM.óLå¯ä˙å¿èIóπ = ""
               HTM.éËîzãÊï™ = ""
               HTM.î[ì¸êÊ = ""
               HTM.î[ì¸êÊéÛì¸ = ""
               HTM.ïΩèÄâªãÊï™ = 0
               HTM.ïΩèÄâªåWêî = 0
               HTM.åWêî1 = 0
               HTM.åWêî2 = 0
               HTM.åWêî3 = 0
               HTM.LT = 0
               HTM.ÉçÉbÉg1 = 0
               HTM.ÉçÉbÉg2 = 0
               HTM.Week1 = 0
               HTM.Week2 = 0
               HTM.Week3 = 0
               HTM.Week4 = 0
               HTM.Week5 = 0
               HTM.Day1 = 0
               HTM.Day2 = 0
               HTM.Day3 = 0
               HTM.Day4 = 0
               HTM.Day5 = 0
               
               HTM.ì‡é¶î≠íç = 0
               HTM.äÓèÄçHêî = 0
               HTM.îıçl = iB3_JBiko
'================================================ *** 2009/07/22 Start
'               HTM.PNO = CisFun.Val2(iB3_JPNo)
               HTM.PNO = 0
'================================================ *** 2009/07/22 End
               
               HTM.ì`ï[éÌóﬁ = iB3_DenSyuJ
                
               HTM.î[ïièëå`ë‘ = RTrim(SaveKeitai)
               HTM.éÛóÃèëå`ë‘ = RTrim(SaveKeitai)
'''2007/12/11DEL               With CisComboJNou
'''2007/12/11DEL                   .CodeGet
'''2007/12/11DEL                   HTM.î[ïièëå`ë‘ = .Code
'''2007/12/11DEL               End With
'''2007/12/11DEL               With CisComboJJyu
'''2007/12/11DEL                   .CodeGet
'''2007/12/11DEL                   HTM.éÛóÃèëå`ë‘ = .Code
'''2007/12/11DEL               End With
               HTM.î[ïièëî≠çsãÊï™ = B3Chk_PrtkbNJ.Value
               HTM.éxããèëî≠çsãÊï™ = B3Chk_PrtkbSJ.Value
               HTM.éÛóÃèëî≠çsãÊï™ = B3Chk_PrtkbJJ.Value
               HTM.éÛóÃèëî≠çs = B3Chk_JuryoJ.Value
               HTM.ñæç◊èëî≠çs = ""
               
               HTM.íºëóì`ï[î≠çsãÊï™ = 0                 '2009/01/19

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
               With CisComboJsya
                   .CodeGet
                   HTM.é©é–çHãÊ = .Code
               End With
               If B3Chk_GouhoJ.Value = 0 Then
                  HTM.çÜï‚ãÊï™ = "G"
               Else
                  HTM.çÜï‚ãÊï™ = "H"
               End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
               
               '====================== *** 2009/07/06 Start
               '''HTM.èoâ◊èÍ = RTrim(iB3_Syukaba)
               '====================== *** 2009/07/06 End
            
               '====================== *** 2009/07/22 Start
               HTM.ïîïiãÊï™ = 0
               HTM.äÓèÄíPâø = 0
               HTM.éwé¶èëî≠çsãÊï™ = B3Chk_Sijisyo.Value
               HTM.èoâ◊ä«óùïîèê = RTrim(iB3_SyukaK)
               '====================== *** 2009/07/22 End
            
            End If
        End With
'-------------------------------------------------------' 2009.10.20 add start
''''    If HTM.éËîzãÊï™ = 2 Then                                                ' 2009.10.23 del
        If B2Op_Hinmoku(1).Value = True Then            ' î≠íçÇ»ÇÁÇŒílÇÉZÉbÉg  ' 2009.10.23 add
           HTM.É|ÉCÉìÉg = RTrim(iB3_Point.Text)
           HTM.ï‘ãpèÍèä = RTrim(iB3_HenCD.Text)
           HTM.å„çHíˆ = CisFun.Val2(RTrim(iB3_AtoCD.Text))
        Else
           HTM.É|ÉCÉìÉg = ""                            ' éÛíçÇ»ÇÁÇŒÉfÉtÉHÉãÉgíl
           HTM.ï‘ãpèÍèä = ""
           HTM.å„çHíˆ = 0
        End If
'-------------------------------------------------------' 2009.10.20 add end
        
        ' çXêVèÓïÒÉOÉäÉbÉhï\é¶
        Call TorInfoToGrid(wRow)
        
        .Redraw = True
    
        Call CisDB.DBTran(TransBegin)  ' ƒ◊›ªﬁ∏ºÆ›äJén
        
        If AddFlg Then
           HTM.çÏê¨é“ = gTanto
           Call HTMInsert
        Else
           HTM.çXêVé“ = gTanto
           Call HTMUpdate
        End If
        ' çÏã∆èÓïÒçXêV
        'If Not DBPutSagyo Then Exit Function                       2009/07/22
        '+--------------------------------------- 2004/12/13 Insert
        ' ã@äBèÓïÒçXêV
        'If Not DBPutKikai Then Exit Function                       2009/07/22
        '+--------------------------------------- 2004/12/13 Insert End
        
            
        Call DBPutYouki(iB3_Youki1)
        Call DBPutYouki(iB3_Youki2)
            
        Call CisDB.DBTran(TransCommit)  ' ƒ◊›ªﬁ∏ºÆ›èIóπ
        
        mTorFlg = True
    End With
    
    DBPutTor = True
    Exit Function
DBPutTor_Err:
    CisFun.ErrorBox
    End
End Function
'+--------------------------------------+
'+      çÏã∆èÓïÒÅ@çXêV
'+--------------------------------------+
Public Function DBPutSagyo() As Boolean
    Dim AddFlg  As Boolean
    DBPutSagyo = False
'====================================================== *** 2009/07/22 Start
' éÊà¯êÊçÏã∆èÓïÒçXêV
'    With vsGridS
'        gSL_Select = "Update ïiî‘éÊà¯çÏã∆É}ÉXÉ^ Set çÌèúãÊï™ = 'D'"
'        gSL_Select = gSL_Select & " Where ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
'        gSL_Select = gSL_Select & "   And éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
'        gSL_Select = gSL_Select & "   And éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
'        CisDB.SQL = gSL_Select
'        Call CisDB.DBExec
'
'        For mIx = 1 To .Rows - 1
'            If Trim(.TextMatrix(mIx, 0)) <> "" Or Trim(.TextMatrix(mIx, 1)) <> "" Then
'                gSL_Select = "Select * from ïiî‘éÊà¯çÏã∆É}ÉXÉ^ "
'                gSL_Select = gSL_Select & " Where ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
'                gSL_Select = gSL_Select & " And   éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
'                gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
'                gSL_Select = gSL_Select & " And   çÏã∆èá = " & CisFun.Val2(.TextMatrix(mIx, 0))
'                If HSMRead(gSL_Select, 1) Then
'                   AddFlg = False
'                Else
'                   AddFlg = True
'                   HSM.ïiî‘ = RTrim(iH1_Hinbn)
'                   HSM.éÊà¯êÊ = RTrim(iB2_Torcd)
'                   HSM.éÛì¸ = RTrim(iB2_Ukeir)
'                   HSM.çÏã∆èá = CisFun.Val2(.TextMatrix(mIx, 0))
'                End If
'
'                HSM.çÏã∆ = Trim(.TextMatrix(mIx, 1))
'                HSM.îıçl = RTrim(.TextMatrix(mIx, 3))
'                HSM.çÌèúãÊï™ = ""
'
'                If AddFlg Then
'                   Call HSMInsert
'                Else
'                   Call HSMUpdate
'                End If
'            End If
'        Next mIx
'        With CisDB
'            .SQL = "delete ïiî‘éÊà¯çÏã∆É}ÉXÉ^ "
'            .SQL = .SQL & " Where çÌèúãÊï™ = 'D'"
'            .DBExec
'        End With
'    End With
'====================================================== *** 2009/07/22 End
    DBPutSagyo = True
End Function
'+---------------------------------------------- 2004/12/13 Insert
'+--------------------------------------+
'+      ã@äBèÓïÒÅ@çXêV
'+--------------------------------------+
Public Function DBPutKikai() As Boolean
    Dim AddFlg  As Boolean
    DBPutKikai = False

'====================================================== *** 2009/07/22 Start
' éÊà¯êÊã@äBèÓïÒçXêV
'    With VsGridK
'        gSL_Select = "Update ïiî‘éÊà¯ã@äBÉ}ÉXÉ^ Set çÌèúãÊï™ = 'D'"
'        gSL_Select = gSL_Select & " Where ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
'        gSL_Select = gSL_Select & "   And éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
'        gSL_Select = gSL_Select & "   And éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
'        CisDB.SQL = gSL_Select
'        Call CisDB.DBExec
'
'        For mIx = 1 To .Rows - 1
'            If Trim(.TextMatrix(mIx, 1)) <> "" Then
'                gSL_Select = "Select * from ïiî‘éÊà¯ã@äBÉ}ÉXÉ^ "
'                gSL_Select = gSL_Select & " Where ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
'                gSL_Select = gSL_Select & " And   éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
'                gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
'                gSL_Select = gSL_Select & " And   ã@äB = '" & Trim(.TextMatrix(mIx, 1)) & "'"
'                If HKMRead(gSL_Select, 1) Then
'                   AddFlg = False
'                Else
'                   AddFlg = True
'                   HKM.ïiî‘ = RTrim(iH1_Hinbn)
'                   HKM.éÊà¯êÊ = RTrim(iB2_Torcd)
'                   HKM.éÛì¸ = RTrim(iB2_Ukeir)
'                   HKM.ã@äB = Trim(.TextMatrix(mIx, 1))
'                End If
'
'                If Trim(.TextMatrix(mIx, 3)) <> "" Then
'                    HKM.äÓñ{ = "1"
'                    HKM.ã@äBîıçl = RTrim(iB3_SJBiko)
'                Else
'                    HKM.äÓñ{ = ""
'                    HKM.ã@äBîıçl = ""
'                End If
'
'                HKM.çÌèúãÊï™ = ""
'
'                If AddFlg Then
'                   Call HKMInsert
'                Else
'                   Call HKMUpdate
'                End If
'            End If
'        Next mIx
'        With CisDB
'            .SQL = "delete ïiî‘éÊà¯ã@äBÉ}ÉXÉ^ "
'            .SQL = .SQL & " Where ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
'            .SQL = .SQL & "   And éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
'            .SQL = .SQL & "   And éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
'            .SQL = .SQL & "   And çÌèúãÊï™ = 'D'"
'            .DBExec
'        End With
'    End With
'====================================================== *** 2009/07/22 End
    DBPutKikai = True
End Function
'+---------------------------------------------- 2004/12/13 Insert End

'+------------------------------+
'+  é˚óeäÌÅ@çXêV
'+------------------------------+
Private Function DBPutYouki(Youki As String) As Boolean
    
    If Trim(Youki) = "" Then Exit Function
        
    gSL_Select = "Select * from ñºèÃÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " Where ãÊï™ñºèÃ = 'é˚óeäÌä«óù'"
    gSL_Select = gSL_Select & " And   ãÊï™É^ÉCÉv = ''"
    gSL_Select = gSL_Select & " And   íl = '" & RTrim(Youki) & "'"
    If SYMRead(gSL_Select, 1, 1) Then Exit Function
    
    Call ItemsClearSYM
    With SYM
        .ãÊï™ñºèÃ = "é˚óeäÌä«óù"
        .íl = RTrim(Youki)
        .ãÊï™åÖêî = 10
        .çÏê¨é“ = gTanto
    End With
    Call SYMInsert

End Function
'+==========================================================+
'+      ÉOÉäÉbÉhèàóù                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   å©èoÇµæØƒ & èâä˙èÛë‘æØƒ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'Å@∏ﬁÿØƒﬁèâä˙âªèàóù
'        .Clear
'       Call .FixedSet(å©èoÇµà íuê›íË,óÒà íuê›íË,œ∞ºﬁèàóù,
'                      æŸåÖêî,óÒî‘çÜ,å©èoÇµì‡óe,ï™óﬁ¿≤Ãﬂ)
        ReDim SCol(45)                                      ' 2009.10.20 upd : ÉTÉCÉY 41-->45
        ' No|éÊà¯êÊ|éÛ|îwî‘çÜ|ïiñ⁄|éËîz|î‰ó¶|SKãÊï™|ª≤∏Ÿ|é˚óeêî1|ïœçXì˙-ï÷|é˚óeêî2|ÉAÉhÉåÉX|å_ñÒêÊ|ì`ï[î≠çs|î[ïièëå`ë‘|óLå¯ä˙å¿|éÛóÃèëå`ë‘|éÛóÃèëî≠çs|ñæç◊èëî≠çs|Ç©åWêî|ïΩèÄâª|åWêî1|åWêî2|åWêî3|LT|€Øƒá@|€ØƒáA|ójì˙|ì˙|çHêî|îıçl|î[ì¸êÊ|î[éÛ|é˚óeäÌ1|é˚óeäÌ2|íuèÍ|ç≈èIî≠çsNO|éÊà¯êÊCD|ïiñ⁄CD|Select
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 0, 0, "No")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 20, 1, "éÊà¯êÊ")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 2, 2, "éÛ")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 3, "îwî‘çÜ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 5, 4, "ïiñ⁄")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 6, 5, "éËîz")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 0, 6, "î‰ó¶")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 8, 7, "SKãÊï™")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 7, 8, "ª≤∏Ÿ")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 0, 9, "é˚óeêî1")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 14, 10, "ïœçXì˙-ï÷")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 0, 11, "é˚óeêî2")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 12, "ÉAÉhÉåÉX")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 20, 13, "å_ñÒêÊ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 14, "ì`ï[î≠çs")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 15, "î[ïièëå`ë‘")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 22, 16, "óLå¯ä˙å¿")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 17, "éÛóÃèëå`ë‘")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 18, "éÛóÃèëî≠çs")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 19, "ñæç◊èëî≠çs")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 20, "Ç©åWêî")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 6, 21, "ïΩèÄâª")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 0, 22, "åWêî1")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 0, 23, "åWêî2")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 0, 24, "åWêî3")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 0, 25, "LT")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 7, 26, "€Øƒá@")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 7, 27, "€ØƒáA")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 14, 28, "ój ì˙ éw íË")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 14, 29, "ì˙ ït éw íË")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 0, 30, "çHêî")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 30, 31, "îıçl")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 20, 32, "î[ì¸êÊ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 33, "î[éÛ")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 10, 34, "é˚óeäÌ1")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 10, 35, "é˚óeäÌ2")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 10, 36, "íuèÍ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 37, "ç≈èIî≠çsNO")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 8, 38, "ì‡é¶î≠íç")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 39, "éÊà¯êÊCD")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 40, "ïiñ⁄CD")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 0, 41, "PNO")
'Å@∏ﬁÿØƒﬁëIëèàóùópÅ@ëﬁî¥ÿ±
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 42, "Select")
'-------------------------------------------------------------------' 2009.10.20 add start
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 43, "É|ÉCÉìÉg")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 44, "ï‘ãpèÍèä")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 45, "å„çHíˆ")
'-------------------------------------------------------------------' 2009.10.20 add end
'  ï\é¶ ﬂ¿∞›æØƒ
'       Call .PatanSet( ﬂ¿∞›NO,ï\é¶ÉpÉ^Å[Éì)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,41,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38")
'  ï“èWì‡óeÉZÉbÉg
'       Call .EditSet(√∑ΩƒµÃﬁºﬁ™∏ƒ,∂◊—NO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
        
        SCol(1) = .FixedGet(1)  ' éÊà¯êÊ
        SCol(2) = .FixedGet(2)  ' éÛì¸
        SCol(3) = .FixedGet(3)  ' îwî‘çÜ
        SCol(4) = .FixedGet(4)  ' ïiñ⁄
        SCol(5) = .FixedGet(5)  ' éËîz
        SCol(6) = .FixedGet(6)  ' î‰ó¶
        SCol(7) = .FixedGet(7)  ' SKãÊï™
        SCol(8) = .FixedGet(8)  ' ª≤∏Ÿ
        SCol(9) = .FixedGet(9)  ' é˚óeêî
        SCol(10) = .FixedGet(10)  ' ïœçXì˙-ï÷
        SCol(11) = .FixedGet(11)  ' é˚óeêî2
        SCol(12) = .FixedGet(12)  ' ÉAÉhÉåÉX
        SCol(13) = .FixedGet(13)  ' å_ñÒêÊ
        SCol(14) = .FixedGet(14)  ' ì`ï[î≠çs
        SCol(15) = .FixedGet(15)  ' î[ïièëå`ë‘
        SCol(16) = .FixedGet(16)  ' óLå¯ä˙å¿
        SCol(17) = .FixedGet(17)  ' éÛóÃèëå`ë‘
        SCol(18) = .FixedGet(18)  ' éÛóÃèëî≠çs
        SCol(19) = .FixedGet(19)  ' ñæç◊èëî≠çs
        SCol(20) = .FixedGet(20)  ' Ç©åWêî
        SCol(21) = .FixedGet(21)  ' ïΩèÄâª
        SCol(22) = .FixedGet(22)  ' åWêî1
        SCol(23) = .FixedGet(23)  ' åWêî2
        SCol(24) = .FixedGet(24)  ' åWêî4
        SCol(25) = .FixedGet(25)  ' LT
        SCol(26) = .FixedGet(26)  ' €Øƒá@
        SCol(27) = .FixedGet(27)  ' €ØƒáA
        SCol(28) = .FixedGet(28)  ' ójì˙
        SCol(29) = .FixedGet(29)  ' ì˙
        SCol(30) = .FixedGet(30)  ' çHêî
        SCol(31) = .FixedGet(31)  ' îıçl
        SCol(32) = .FixedGet(32)  ' î[ì¸êÊ
        SCol(33) = .FixedGet(33)  ' î[éÛ
        SCol(34) = .FixedGet(34)  ' é˚óeäÌ1
        SCol(35) = .FixedGet(35)  ' é˚óeäÌ2
        SCol(36) = .FixedGet(36)  ' íuèÍ
        SCol(37) = .FixedGet(37)  ' ç≈èIî≠çsNO
        SCol(38) = .FixedGet(38)  ' ì‡é¶î≠íç
        SCol(39) = .FixedGet(39)  ' éÊà¯êÊCD
        SCol(40) = .FixedGet(40)  ' ïiñ⁄CD
        SCol(41) = .FixedGet(41)  ' PNO
        SCol(43) = .FixedGet(43)  ' É|ÉCÉìÉg    ' 2009.10.20 add
        SCol(44) = .FixedGet(44)  ' ï‘ãpèÍèä    ' 2009.10.20 add
        SCol(45) = .FixedGet(45)  ' å„çHíˆ      ' 2009.10.20 add
    End With
    KCol = 39   ' îÒï\é¶çÄñ⁄äJén∂◊—
    
End Sub
'***********************************************'
'*****       ÉOÉäÉbÉhì‡óeÉZÉbÉg(éÊà¯êÊèÓïÒ)  *****
'***********************************************'
Private Sub GridSet()
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
' íäèoèàóù èâä˙æØƒ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' íäèoåèêî(çáåv)ëﬁî
' ∏ﬁÿØƒﬁçsçƒæØƒ & óÒà íuéÊìæ
    With vsGrid1
       .Redraw = False
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not HTM_RDSTS
        ID = ID + 1
            
        Call TorInfoToGrid(ID)
        
        Call HTMReadNext(1)
    Loop
    With vsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'+-------------------------------------+
'+  ïiî‘éÊà¯êÊèÓïÒÇÊÇËÉOÉäÉbÉhÇ÷ÉZÉbÉg    +
'+-------------------------------------+
Private Sub TorInfoToGrid(SetRow As Long)
    With vsGrid1
        For mIx = 0 To .Cols - 1
            .TextMatrix(SetRow, mIx) = ""
        Next mIx
        
        ' No
        .TextMatrix(SetRow, 0) = SetRow
        'éÊà¯êÊ
        If HTM.ïiñ⁄ = 0 Then gInt = 3 Else gInt = 9
        If TorNmGet(HTM.éÊà¯êÊ, gInt) Then
           .TextMatrix(SetRow, SCol(1)) = HTM.éÊà¯êÊ & ":" & TRM_RName
        Else
           .TextMatrix(SetRow, SCol(1)) = HTM.éÊà¯êÊ
        End If
        .TextMatrix(SetRow, SCol(39)) = HTM.éÊà¯êÊ
        ' éÛì¸
        .TextMatrix(SetRow, SCol(2)) = HTM.éÛì¸
        ' îwî‘çÜ
        .TextMatrix(SetRow, SCol(3)) = HTM.îwî‘çÜ
        ' ïiñ⁄
        .TextMatrix(SetRow, SCol(4)) = B2Op_Hinmoku(HTM.ïiñ⁄).Caption
        .TextMatrix(SetRow, SCol(40)) = HTM.ïiñ⁄
        ' éËîzãÊï™
        With CisComboTehai
            .Code = HTM.éËîzãÊï™
            .NameGet
            vsGrid1.TextMatrix(SetRow, SCol(5)) = .CombText
            vsGrid1.TextMatrix(SetRow, SCol(35)) = .Code
        End With
        ' î‰ó¶
        If HTM.î≠íçî‰ó¶ <> 0 Then
           .TextMatrix(SetRow, SCol(6)) = HTM.î≠íçî‰ó¶ & "%"
        End If
        ' SK
        If HTM.SKãÊï™ = 0 Then
           .TextMatrix(SetRow, SCol(7)) = "Ç©ÇÒÇŒÇÒ"
        Else
           .TextMatrix(SetRow, SCol(7)) = "éwé¶"
        End If
        ' ÉTÉCÉNÉã
        If Trim(HTM.ÉTÉCÉNÉã) <> "" Then
            .TextMatrix(SetRow, SCol(8)) = Mid(HTM.ÉTÉCÉNÉã, 1, 1) & "-" & _
                                     Mid(HTM.ÉTÉCÉNÉã, 2, 2) & "-" & _
                                     Mid(HTM.ÉTÉCÉNÉã, 4, 2)
        End If
        ' é˚óeêî1
        .TextMatrix(SetRow, SCol(9)) = Format(HTM.é˚óeêî1, "#,###")
        ' ïœçXì˙Å[ï÷
        If Trim(HTM.ïœçXì˙) <> "" Then
           .TextMatrix(SetRow, SCol(10)) = Mid(HTM.ïœçXì˙, 1, 4) & "/" & _
                                     Mid(HTM.ïœçXì˙, 5, 2) & "/" & _
                                     Mid(HTM.ïœçXì˙, 7, 2) & "-" & _
                                     Format(CisFun.Val2(HTM.ïœçXï÷), "00")
        End If
        ' é˚óeêî2
        .TextMatrix(SetRow, SCol(11)) = Format(HTM.é˚óeêî2, "#,###")
        ' ÉAÉhÉåÉX
        .TextMatrix(SetRow, SCol(12)) = Format(HTM.ÉAÉhÉåÉX, iB3_Adress.cFormat)
'        ' å_ñÒêÊ
        If HTM.ïiñ⁄ = 0 Then gInt = 2 Else gInt = 5
        If TorNmGet(HTM.å_ñÒêÊ, gInt) Then
           .TextMatrix(SetRow, SCol(13)) = HTM.å_ñÒêÊ & ":" & TRM_RName
        Else
           .TextMatrix(SetRow, SCol(13)) = HTM.å_ñÒêÊ
        End If
        ' óLå¯ä˙å¿
        gStr = ""
        If Trim(HTM.óLå¯ä˙å¿äJén) <> "" Then
            gStr = Mid(HTM.óLå¯ä˙å¿äJén, 1, 4) & "/" & _
                   Mid(HTM.óLå¯ä˙å¿äJén, 5, 2) & "/" & _
                   Mid(HTM.óLå¯ä˙å¿äJén, 7, 2)
        End If
        If Trim(HTM.óLå¯ä˙å¿èIóπ) <> "" Then
           If gStr = "" Then gStr = Space(10)
           gStr = gStr & "Å`"
           gStr = gStr & Mid(HTM.óLå¯ä˙å¿èIóπ, 1, 4) & "/" & _
                         Mid(HTM.óLå¯ä˙å¿èIóπ, 5, 2) & "/" & _
                         Mid(HTM.óLå¯ä˙å¿èIóπ, 7, 2)
        Else
           If Trim(HTM.óLå¯ä˙å¿äJén) <> "" Then
              gStr = gStr & "Å`" & Space(10)
           End If
        End If
        .TextMatrix(SetRow, SCol(16)) = gStr
'
        ' Ç©åWêî
        If HTM.Ç©åWêî <> 0 Then
           .TextMatrix(SetRow, SCol(20)) = Format(HTM.Ç©åWêî, iB3_KKeisu.cFormat)
        End If

        If HTM.ïiñ⁄ = 0 Then
            ' ì`ï[éÌóﬁ
            Call MeisyoGet("ì`ï[éÌóﬁéÛíç", HTM.ì`ï[éÌóﬁ)
            .TextMatrix(SetRow, SCol(14)) = SYM_Meisyo
            ' î[ïièëå`ë‘
            Call MeisyoGet("î[ïièëå`ë‘éÛíç", HTM.î[ïièëå`ë‘)
            .TextMatrix(SetRow, SCol(15)) = SYM_Meisyo
            ' éÛóÃèëå`ë‘
            Call MeisyoGet("éÛóÃèëå`ë‘éÛíç", HTM.éÛóÃèëå`ë‘)
            .TextMatrix(SetRow, SCol(17)) = SYM_Meisyo
            ' éÛóÃèëî≠çs
            If HTM.éÛóÃèëå`ë‘ <> "" Then
                If HTM.éÛóÃèëî≠çs = 1 Then
                   .TextMatrix(SetRow, SCol(18)) = "èoâ◊éû"
                Else
                   .TextMatrix(SetRow, SCol(18)) = "éÛíçéû"
                End If
            End If
        Else
            ' ì`ï[éÌóﬁ
            Call MeisyoGet("ì`ï[éÌóﬁî≠íç", HTM.ì`ï[éÌóﬁ)
            .TextMatrix(SetRow, SCol(14)) = SYM_Meisyo
            ' î[ïièëå`ë‘
            Call MeisyoGet("î[ïièëå`ë‘î≠íç", HTM.î[ïièëå`ë‘)
            .TextMatrix(SetRow, SCol(15)) = SYM_Meisyo
            ' éÛóÃèëå`ë‘
            Call MeisyoGet("éÛóÃèëå`ë‘", HTM.éÛóÃèëå`ë‘)
            .TextMatrix(SetRow, SCol(17)) = SYM_Meisyo
            ' ñæç◊èëî≠çs
            Call MeisyoGet("ñæç◊èëî≠çs", HTM.ñæç◊èëî≠çs)
            .TextMatrix(SetRow, SCol(19)) = SYM_Meisyo
            ' éÛóÃèëî≠çs
            If HTM.éÛóÃèëå`ë‘ <> "" Then
                If HTM.éÛóÃèëî≠çs = 1 Then
                   .TextMatrix(SetRow, SCol(18)) = "éÛóÃéû"
                Else
                   .TextMatrix(SetRow, SCol(18)) = "î≠íçéû"
                End If
            End If
            ' ïΩèÄâªãÊï™
            gLong = HTM.ïΩèÄâªãÊï™
            If gLong <> 0 And HTM.ïΩèÄâªåWêî <> 0 Then
               .TextMatrix(SetRow, SCol(21)) = B3Op_Heijun(gLong).Caption & " " & HTM.ïΩèÄâªåWêî & "%"
            Else
               .TextMatrix(SetRow, SCol(21)) = B3Op_Heijun(gLong).Caption
            End If
        End If
        ' åWêî1
        If HTM.åWêî1 <> 0 Then
           .TextMatrix(SetRow, SCol(22)) = HTM.åWêî1
        End If
        ' åWêî2
        If HTM.åWêî2 <> 0 Then
           .TextMatrix(SetRow, SCol(23)) = HTM.åWêî2
        End If
        ' åWêî3
        If HTM.åWêî3 <> 0 Then
           .TextMatrix(SetRow, SCol(24)) = HTM.åWêî3
        End If
        ' LT
        If HTM.LT <> 0 Then
           .TextMatrix(SetRow, SCol(25)) = HTM.LT
        End If
        ' ÉçÉbÉgá@
        If HTM.ÉçÉbÉg1 <> 0 Then
           .TextMatrix(SetRow, SCol(26)) = Format(HTM.ÉçÉbÉg1, iB3_LSu.cFormat)
        End If
        ' ÉçÉbÉgáA
        If HTM.ÉçÉbÉg2 <> 0 Then
           .TextMatrix(SetRow, SCol(27)) = Format(HTM.ÉçÉbÉg2, iB3_LSu.cFormat)
        End If
        ' ójì˙
        gStr = "": mStr = "•"
        If HTM.Week1 <> 0 Then
           gStr = B3Chk_Week(0).Caption
        Else
           gStr = "Å@"
        End If
        If HTM.Week2 <> 0 Then
           gStr = gStr & mStr & B3Chk_Week(1).Caption
        Else
           gStr = gStr & mStr & "Å@"
        End If
        If HTM.Week3 <> 0 Then
           gStr = gStr & mStr & B3Chk_Week(2).Caption
        Else
           gStr = gStr & mStr & "Å@"
        End If
        If HTM.Week4 <> 0 Then
           gStr = gStr & mStr & B3Chk_Week(3).Caption
        Else
           gStr = gStr & mStr & "Å@"
        End If
        If HTM.Week5 <> 0 Then
           gStr = gStr & mStr & B3Chk_Week(4).Caption
        Else
           gStr = gStr & mStr & "Å@"
        End If
        If gStr <> "" Then
            gStr = Replace(gStr, "Å@•Å@•Å@•Å@•Å@", "")
           .TextMatrix(SetRow, SCol(28)) = gStr
        End If
        ' ì˙
        gStr = ""
        If HTM.Day1 <> 0 Then
           gStr = CisFun.RSetFld(HTM.Day1, 0, iB3_Day(0).cFormat)
        End If
        If HTM.Day2 <> 0 Then
           If gStr <> "" Then gStr = gStr & "•"
           gStr = gStr & CisFun.RSetFld(HTM.Day2, 0, iB3_Day(1).cFormat)
        End If
        If HTM.Day3 <> 0 Then
           If gStr <> "" Then gStr = gStr & "•"
           gStr = gStr & CisFun.RSetFld(HTM.Day3, 0, iB3_Day(2).cFormat)
        End If
        If HTM.Day4 <> 0 Then
           If gStr <> "" Then gStr = gStr & "•"
           gStr = gStr & CisFun.RSetFld(HTM.Day4, 0, iB3_Day(3).cFormat)
        End If
        If HTM.Day5 <> 0 Then
           If gStr <> "" Then gStr = gStr & "•"
           gStr = gStr & CisFun.RSetFld(HTM.Day5, 0, iB3_Day(4).cFormat)
        End If
        If gStr <> "" Then
           .TextMatrix(SetRow, SCol(29)) = gStr
        End If
'        ' ì‡é¶î≠íç
        If HTM.ì‡é¶î≠íç = 1 Then
           .TextMatrix(SetRow, SCol(38)) = "Åõ"
        End If
        ' çHêî
        If HTM.äÓèÄçHêî <> 0 Then
           .TextMatrix(SetRow, SCol(30)) = Format(HTM.äÓèÄçHêî, iB3_Kosu.cFormat)
        End If
        ' îıçl
         .TextMatrix(SetRow, SCol(31)) = HTM.îıçl
        ' î[ì¸êÊ
        If HTM.éËîzãÊï™ = "1" Then
            If HTM.ïiñ⁄ = 0 Then gInt = 3 Else gInt = 8
            If TorNmGet(HTM.î[ì¸êÊ, gInt) Then
               .TextMatrix(SetRow, SCol(32)) = HTM.î[ì¸êÊ & ":" & TRM_RName
            Else
               .TextMatrix(SetRow, SCol(32)) = HTM.î[ì¸êÊ
            End If
            ' î[ì¸êÊéÛì¸
            .TextMatrix(SetRow, SCol(33)) = HTM.î[ì¸êÊéÛì¸
        End If
        ' é˚óeäÌ1
        .TextMatrix(SetRow, SCol(34)) = HTM.é˚óeäÌ1
        ' é˚óeäÌ2
        .TextMatrix(SetRow, SCol(35)) = HTM.é˚óeäÌ2
        ' íuèÍ
        .TextMatrix(SetRow, SCol(36)) = HTM.íuèÍ
        ' ç≈èIî≠çsNo
        If HTM.ç≈èIî≠çsòAî‘ <> 0 Then
           .TextMatrix(SetRow, SCol(37)) = Format(HTM.ç≈èIî≠çsòAî‘, "000")
        End If
        ' PNO
        If HTM.PNO <> 0 Then
           .TextMatrix(SetRow, SCol(41)) = Format(HTM.PNO, "##")
        End If
'-------------------------------------------------------------------' 2009.10.20 add start
        ' É|ÉCÉìÉg/ï‘ãpèÍèä/å„çHíˆ
        .TextMatrix(SetRow, SCol(43)) = HTM.É|ÉCÉìÉg
        .TextMatrix(SetRow, SCol(44)) = HTM.ï‘ãpèÍèä
        If HTM.å„çHíˆ <> 0 Then
           .TextMatrix(SetRow, SCol(45)) = Format(HTM.å„çHíˆ, "000")
        Else
           .TextMatrix(SetRow, SCol(45)) = ""
        End If
'-------------------------------------------------------------------' 2009.10.20 add end
    End With
End Sub
'*************************************************
'*      éÊà¯êÊèÓïÒÉOÉäÉbÉhÇÊÇËì¸óÕàÊÇ÷ÉZÉbÉg        *
'*************************************************
Private Sub GridToText(Optional SetRow As Long = 0)
    If SetRow = 0 Then
       mIx = CisFun.Val2(iB2_Torcd.Tag)
    Else
       mIx = SetRow
    End If
    
    With vsGrid1
        gSL_Select = "Select HT.*,TU.ójì˙î≠íç from ïiî‘éÊà¯êÊÉ}ÉXÉ^ HT"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN éÊà¯êÊéÛì¸É}ÉXÉ^ TU"
        gSL_Select = gSL_Select & "   ON TU.éÊà¯êÊCD = HT.éÊà¯êÊ"
        gSL_Select = gSL_Select & "   AND TU.éÛì¸ = HT.éÛì¸"
        gSL_Select = gSL_Select & "   AND TU.éÊà¯êÊãÊï™ = HT.ïiñ⁄"
        gSL_Select = gSL_Select & " Where HT.ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " And   HT.éÊà¯êÊ = '" & Trim(.TextMatrix(mIx, SCol(39))) & "'"
        gSL_Select = gSL_Select & " And   HT.éÛì¸ = '" & Trim(.TextMatrix(mIx, SCol(2))) & "'"
        If Not HTMRead(gSL_Select, 1) Then
           Exit Sub
        End If
        If SetRow = 0 Then
            ' éÊà¯êÊ
            iB2_Torcd = HTM.éÊà¯êÊ
            Call iB2_Torcd_LostFocus
            ' éÛì¸
            iB2_Ukeir = HTM.éÛì¸
            ' ïiñ⁄
            B2Op_Hinmoku(0).Tag = HTM.ïiñ⁄
            If HTM.ïiñ⁄ = 0 Then
               B2Op_Hinmoku(0).Value = True
            Else
               B2Op_Hinmoku(1).Value = True
            End If
            ' ç≈èIî≠çsNo
            If HTM.ç≈èIî≠çsòAî‘ <> 0 Then
               B3lb_HNo = Format(HTM.ç≈èIî≠çsòAî‘, "000")
            End If
        End If
'        ' îwî‘çÜ
        iB3_Seban = HTM.îwî‘çÜ
        ' éËîzãÊï™
        With CisComboTehai
            .Code = HTM.éËîzãÊï™
            .TextGet
'*--------------------------------------------------<< 2004/06/28 Delete Start >>
'           If .Code = "2" Then
'              B3lb_Tehai = .CombText
'              B3lb_Tehai.Visible = True
'              B3Comb_Tehai.Visible = False
'           Else
'              B3lb_Tehai.Visible = False
'              B3Comb_Tehai.Visible = True
'           End If
            TRM.é–ì‡ãÊï™ = 0
            gSL_Select = ""
            gSL_Select = gSL_Select & "select é–ì‡ãÊï™"
            gSL_Select = gSL_Select & "  from éÊà¯êÊÉ}ÉXÉ^ "
            gSL_Select = gSL_Select & " where éÊà¯êÊãÊï™ = " & HTM.ïiñ⁄ & ""
            gSL_Select = gSL_Select & "   and éÊà¯êÊCD = '" & HTM.éÊà¯êÊ & "'"
            Call TRMRead(gSL_Select, 1)
            If TRM.é–ì‡ãÊï™ = 1 Then
                B3lb_Tehai = .CombText
                B3lb_Tehai.Visible = True
                B3Comb_Tehai.Visible = False
            Else
                B3lb_Tehai.Visible = False
                B3Comb_Tehai.Visible = True
            End If
            B3Comb_Tehai.Tag = TRM.é–ì‡ãÊï™
'*--------------------------------------------------<< 2004/06/28 Delete End >>
        End With
        Call B3Comb_Tehai_Click
        
        ' î‰ó¶
        If HTM.î≠íçî‰ó¶ <> 0 Then
            iB3_Hiri = CisFun.RSetFld(HTM.î≠íçî‰ó¶, 0, iB3_Hiri.cFormat)
        End If
        ' SK
        If HTM.SKãÊï™ = 0 Then
           B3Op_SK(0).Value = True
        Else
           B3Op_SK(1).Value = True
        End If
        ' ÉTÉCÉNÉã
        If HTM.ÉTÉCÉNÉã <> "" Then
           iB3_Cycle1 = Mid(HTM.ÉTÉCÉNÉã, 1, 1)
           iB3_Cycle2 = Mid(HTM.ÉTÉCÉNÉã, 2, 2)
           iB3_Cycle3 = Mid(HTM.ÉTÉCÉNÉã, 4, 2)
        End If
        ' é˚óeêî1
        iB3_Syuyo1 = CisFun.RSetFld(HTM.é˚óeêî1, 0, iB3_Syuyo1.cFormat)
        ' ïœçXì˙Å[ï÷
        gStr = Trim(.TextMatrix(mIx, SCol(10)))
        If Trim(HTM.ïœçXì˙) <> "" Then
           iB3_HYmd = Mid(HTM.ïœçXì˙, 1, 8)
           iB3_HBin = Format(HTM.ïœçXï÷, "00")
        End If
        ' é˚óeêî2
        iB3_Syuyo2 = CisFun.RSetFld(HTM.é˚óeêî2, 0, iB3_Syuyo2.cFormat)
        ' ÉAÉhÉåÉX
        iB3_Adress = HTM.ÉAÉhÉåÉX
        ' å_ñÒêÊ
        iB3_KTorcd = HTM.å_ñÒêÊ
        Call iB3_KTorcd_LostFocus
        ' óLå¯ä˙å¿(äJén)
        iB3_YSYmd = HTM.óLå¯ä˙å¿äJén
        ' óLå¯ä˙å¿(èIóπ)
        iB3_YEYmd = HTM.óLå¯ä˙å¿èIóπ
        ' Ç©åWêî
        If HTM.Ç©åWêî <> 0 Then
            iB3_KKeisu = CisFun.RSetFld(HTM.Ç©åWêî, 0, iB3_KKeisu.cFormat)
        End If
        ' ïΩèÄâªãÊï™
        gLong = CisFun.Val2(.TextMatrix(mIx, SCol(41)))

        B3Op_Heijun(0).Tag = HTM.ïΩèÄâªãÊï™
        B3Op_Heijun(HTM.ïΩèÄâªãÊï™).Value = True
        ' ïΩèÄâªåWêî
        If HTM.ïΩèÄâªåWêî <> 0 Then
           iB3_HKeisu = CisFun.RSetFld(HTM.ïΩèÄâªåWêî, 0, iB3_HKeisu.cFormat)
        End If
        ' åWêî1
        If HTM.åWêî1 <> 0 Then
           iB3_Keisu1 = CisFun.RSetFld(HTM.åWêî1, 0, iB3_Keisu1.cFormat)
        End If
        ' åWêî2
        If HTM.åWêî2 <> 0 Then
           iB3_Keisu2 = CisFun.RSetFld(HTM.åWêî2, 0, iB3_Keisu2.cFormat)
        End If
        ' åWêî3
        If HTM.åWêî3 <> 0 Then
           iB3_Keisu3 = CisFun.RSetFld(HTM.åWêî3, 0, iB3_Keisu3.cFormat)
        End If
        ' LT
        If HTM.LT <> 0 Then
           iB3_LT = CisFun.RSetFld(HTM.LT, 0, iB3_LT.cFormat)
        End If
        ' ÉçÉbÉgá@
        If HTM.ÉçÉbÉg1 <> 0 Then
           iB3_LSu = CisFun.RSetFld(HTM.ÉçÉbÉg1, 0, iB3_LSu.cFormat)
        End If
        ' ÉçÉbÉgáA
        If HTM.ÉçÉbÉg2 <> 0 Then
           iB3_LSu2 = CisFun.RSetFld(HTM.ÉçÉbÉg2, 0, iB3_LSu2.cFormat)
        End If
        ' ójì˙
        B3Chk_Week(0).Value = HTM.Week1
        B3Chk_Week(1).Value = HTM.Week2
        B3Chk_Week(2).Value = HTM.Week3
        B3Chk_Week(3).Value = HTM.Week4
        B3Chk_Week(4).Value = HTM.Week5
'        ' ì˙
        If HTM.Day1 <> 0 Then
           iB3_Day(0) = CisFun.RSetFld(HTM.Day1, 0, iB3_Day(0).cFormat)
        End If
        If HTM.Day2 <> 0 Then
           iB3_Day(1) = CisFun.RSetFld(HTM.Day2, 0, iB3_Day(1).cFormat)
        End If
        If HTM.Day3 <> 0 Then
           iB3_Day(2) = CisFun.RSetFld(HTM.Day3, 0, iB3_Day(2).cFormat)
        End If
        If HTM.Day4 <> 0 Then
           iB3_Day(3) = CisFun.RSetFld(HTM.Day4, 0, iB3_Day(3).cFormat)
        End If
        If HTM.Day5 <> 0 Then
           iB3_Day(4) = CisFun.RSetFld(HTM.Day5, 0, iB3_Day(4).cFormat)
        End If
'       ójì˙î≠íç    2005.07.15 í«â¡
        If HTM.ójì˙î≠íç = 0 Then
            Back_Hachu(24).Enabled = True
            Back_Hachu(30).Enabled = True
            Back_Hachu(24).ForeColor = &HFF0000
            Back_Hachu(30).ForeColor = &HFF0000
            B3lb_YoubiHC.Visible = False
        Else
            Back_Hachu(24).Enabled = False
            Back_Hachu(30).Enabled = False
            Back_Hachu(24).ForeColor = &H80000008
            Back_Hachu(30).ForeColor = &H80000008
            B3lb_YoubiHC = "î≠íçóL"
            B3lb_YoubiHC.Visible = True
        End If
'        ' ì‡é¶î≠íç
        B3Chk_Naiji.Value = HTM.ì‡é¶î≠íç
        
        If HTM.ïiñ⁄ = 1 Then
           ' ì`ï[éÌóﬁ
           iB3_DenSyu = HTM.ì`ï[éÌóﬁ
           Call iB3_DenSyu_LostFocus
'''2007/12/11DEL           ' î[ïièëå`ë‘
'''2007/12/11DEL           With CisComboHNou
'''2007/12/11DEL                .Code = HTM.î[ïièëå`ë‘
'''2007/12/11DEL                .TextGet
'''2007/12/11DEL           End With
'''2007/12/11DEL           ' éÛóÃèëå`ë‘
'''2007/12/11DEL           With CisComboHJyu
'''2007/12/11DEL                .Code = HTM.éÛóÃèëå`ë‘
'''2007/12/11DEL                .TextGet
'''2007/12/11DEL           End With
           ' ñæç◊èëî≠çs
           With CisComboHMei
                .Code = HTM.ñæç◊èëî≠çs
                .TextGet
           End With
           '-----* 2008/06/09 Start
           'B3Chk_PrtkbN.Value = HTM.î[ïièëî≠çsãÊï™
           If HTM.éËîzãÊï™ <> 4 Then
              B3Chk_PrtkbN.Value = HTM.î[ïièëî≠çsãÊï™
           Else
              B3Chk_PrtkbSN.Value = HTM.î[ïièëî≠çsãÊï™
           End If
           '-----* 2008/06/09 End
           B3Chk_PrtkbS.Value = HTM.éxããèëî≠çsãÊï™
           B3Chk_PrtkbJ.Value = HTM.éÛóÃèëî≠çsãÊï™
           B3Chk_Juryo.Value = HTM.éÛóÃèëî≠çs
           
            '============================================= 2009/01/15 Start
            If HTM.éËîzãÊï™ = 1 Then
               B3Chk_PrtkbSN1.Value = HTM.íºëóì`ï[î≠çsãÊï™
            End If
            '============================================= 2009/01/15 End

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
           If HTM.çÜï‚ãÊï™ = "G" Then
              B3Chk_Gouho.Value = 0
           Else
              B3Chk_Gouho.Value = 1
           End If
           B3Chk_MaeHak.Value = HTM.ëOçHíˆãÊï™
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
            
           '================================ *** 2009/07/06 Start
           '''iB3_Syukaba = ""
            iB3_SyukaK = ""
           '================================ *** 2009/07/06 End
        Else
           ' ì`ï[éÌóﬁ
           iB3_DenSyuJ = HTM.ì`ï[éÌóﬁ
           Call iB3_DenSyuJ_LostFocus
'''2007/12/11DEL           ' î[ïièëå`ë‘
'''2007/12/11DEL           With CisComboJNou
'''2007/12/11DEL                .Code = HTM.î[ïièëå`ë‘
'''2007/12/11DEL                .TextGet
'''2007/12/11DEL           End With
'''2007/12/11DEL           ' éÛóÃèëå`ë‘
'''2007/12/11DEL           With CisComboJJyu
'''2007/12/11DEL                .Code = HTM.éÛóÃèëå`ë‘
'''2007/12/11DEL                .TextGet
'''2007/12/11DEL           End With
           B3Chk_PrtkbNJ.Value = HTM.î[ïièëî≠çsãÊï™
           B3Chk_PrtkbSJ.Value = HTM.éxããèëî≠çsãÊï™
           B3Chk_PrtkbJJ.Value = HTM.éÛóÃèëî≠çsãÊï™
           B3Chk_JuryoJ.Value = HTM.éÛóÃèëî≠çs
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
           With CisComboJsya
                .Code = HTM.é©é–çHãÊ
                .TextGet
           End With
           If HTM.çÜï‚ãÊï™ = "G" Then
              B3Chk_GouhoJ.Value = 0
           Else
              B3Chk_GouhoJ.Value = 1
           End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
            '================================ *** 2009/07/06 Start
            '''iB3_Syukaba = RTrim(HTM.èoâ◊èÍ)
            '================================ *** 2009/07/06 End
            '================================ *** 2009/07/22 Start
            B3Chk_Sijisyo = HTM.éwé¶èëî≠çsãÊï™
            iB3_SyukaK = RTrim(HTM.èoâ◊ä«óùïîèê)
            Call iB3_SyukaK_LostFocus
            '================================ *** 2009/07/22 End
        End If
        ' çHêî
        If HTM.äÓèÄçHêî <> 0 Then
           iB3_Kosu = CisFun.RSetFld(HTM.äÓèÄçHêî, 0, iB3_Kosu.cFormat)
        End If
        ' îıçl
        iB3_Biko = HTM.îıçl
        iB3_JBiko = iB3_Biko
        ' ŒﬂπØƒáÇ
        iB3_PNo = CisFun.RSetFld(HTM.PNO, 0, iB3_PNo.cFormat)
'        iB3_JPNo = CisFun.RSetFld(HTM.PNO, 0, iB3_PNo.cFormat)             '2009/07/22
        ' î[ì¸êÊ
        iB3_NTorcd = HTM.î[ì¸êÊ
        Call iB3_NTorcd_LostFocus
        ' î[ì¸êÊéÛì¸
        iB3_Ukeir = HTM.î[ì¸êÊéÛì¸
        
        ' ë„çs
        ChkB3_Daiko = HTM.ë„çsãÊï™
        Call ChkB3_Daiko_Click
        iB3_DaikoK = HTM.ë„çså_ñÒêÊ
        Call iB3_DaikoK_LostFocus
        
        ' é˚óeäÌ1
        iB3_Youki1 = HTM.é˚óeäÌ1
        ' é˚óeäÌ2
        iB3_Youki2 = HTM.é˚óeäÌ2
        ' íuèÍ
        iB3_Okiba = HTM.íuèÍ
'---------------------------------------------------------------------------' 2009.10.20 add start
        ' É|ÉCÉìÉg/ï‘ãpèÍèä/å„çHíˆ
        If B2Op_Hinmoku(1) Then                     ' 2009.10.23 add : î≠íçéûÇÃÇ›ÉZÉbÉgÇ∑ÇÈÇÊÇ§èåèÇì¸ÇÍÇΩ
           iB3_Point = HTM.É|ÉCÉìÉg
           iB3_HenCD = HTM.ï‘ãpèÍèä
           If HTM.å„çHíˆ <> 0 Then
              iB3_AtoCD = Format(HTM.å„çHíˆ, "000")
           Else
              iB3_AtoCD = ""
           End If
        Else
           iB3_Point = ""
           iB3_HenCD = ""
           iB3_AtoCD = ""
        End If
        Call iB3_Point_LostFocus
        Call iB3_HenCD_LostFocus
        Call iB3_AtoCD_LostFocus
'---------------------------------------------------------------------------' 2009.10.20 add end
        
        '===================================== *** 2009/07/22 Start
        iB3_Kijyun = Format(HTM.äÓèÄíPâø, "##,##0.00")
        B3Op_HinKu(HTM.ïîïiãÊï™).Value = True
        '===================================== *** 2009/07/22 End
    End With

'---------------------------------------------------------------------------' 2009.10.20 add start
''''If HTM.éËîzãÊï™ = "2" Then  ' éËîzãÊï™Ç™é–ì‡Ç»ÇÁÇŒ                              ' 2009.10.23 del
    If B2Op_Hinmoku(1).Value = True Then  ' éÊà¯êÊãÊï™Ç™î≠íçÇ»ÇÁÇŒ                  ' 2009.10.23 add
       Call SyanaiTehai(True)   ' É|ÉCÉìÉg/ï‘ãpèÍèä/å„çHíˆÅ@çÄñ⁄Çï\é¶
    Else
       Call SyanaiTehai(False)
    End If
'---------------------------------------------------------------------------' 2009.10.20 add end

' çÏã∆èÓïÒÅ@ÉZÉbÉg
    Call SagyoToGrid
'+------------------------------------------------------ 2004/12/13 Insert
' ã@äBèÓïÒÅ@ÉZÉbÉg
    Call KikaiToGrid
'+------------------------------------------------------ 2004/12/13 Insert End

'========================================================== *** 2009/07/22 Start
    Call Kojyun_SQL_Set
    If HKTRead(gSL_Select, , 1) Then
''''       Call GridSet3                ' 2009.10.16 del
      'SAVE_Select = gSL_Select         ' 2009.10.16 add
    End If
    
    If HKTRead(gSL_Select, , 1) Then
       Call GridSet2
    End If
'========================================================== *** 2009/07/22 End
End Sub
'******************************************'
'*****    çÏã∆èÓïÒÉOÉäÉbÉhì‡óeÉZÉbÉg    *****
'******************************************'
Private Sub SagyoToGrid()
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long
    
    gSL_Select = "Select * from ïiî‘éÊà¯çÏã∆É}ÉXÉ^ "
    gSL_Select = gSL_Select & " Where ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
    gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
    gSL_Select = gSL_Select & " Order By çÏã∆èá"
    If Not HSMRead(gSL_Select) Then
       Call HSMClose
       Exit Sub
    End If
    
' íäèoèàóù èâä˙æØƒ
    Ttl_Kensu = CisDB.RecordCount       ' íäèoåèêî(çáåv)ëﬁî
' ∏ﬁÿØƒﬁçsçƒæØƒ & óÒà íuéÊìæ
    '======================================== 20069/07/22 Start
    'With vsGridS
    '   .Redraw = False
    '    If Ttl_Kensu + 1 < CisVsGridS.DispMax Then
    '        .Rows = CisVsGridS.DispMax
    '    Else
    '        .Rows = Ttl_Kensu + 1
    '    End If
    'End With
    'ID = 0
    'Do Until Not HSM_RDSTS
    '    With vsGridS
    '        ID = ID + 1
    '        'áÇ
    '        .TextMatrix(ID, 0) = Format(HSM.çÏã∆èá, "##")
    '        'çÏã∆
    '        .TextMatrix(ID, 1) = HSM.çÏã∆
    '        'îıçl
    '        .TextMatrix(ID, 3) = HSM.îıçl
    '    End With
    '    Call HSMReadNext
    'Loop
    'Call HSMClose
    
    'With vsGridS
    '    For ID = 1 To .Rows - 1
    '        If Trim(.TextMatrix(ID, 1)) = "" Then Exit For
    '        If MeisyoGet("çÏã∆ä«óù", .TextMatrix(ID, 1), 1) Then
    '            .TextMatrix(ID, 2) = SYM_Meisyo
    '        End If
    '    Next ID
    '
    '     .Row = 1: .Col = 1: .ColSel = .Cols - 1
    '     .Redraw = True
    'End With

    '======================================== 20069/07/22 End
End Sub
'+-------------------------------------------------------- 2004/12/13 Insert
'******************************************'
'*****    ã@äBèÓïÒÉOÉäÉbÉhì‡óeÉZÉbÉg    *****
'******************************************'
Private Sub KikaiToGrid()
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long
    
    gSL_Select = "Select * from ïiî‘éÊà¯ã@äBÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " Where ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
    gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
    gSL_Select = gSL_Select & " Order By ã@äB"
    If Not HKMRead(gSL_Select) Then
       Call HKMClose
       Exit Sub
    End If
    
' íäèoèàóù èâä˙æØƒ
    Ttl_Kensu = CisDB.RecordCount       ' íäèoåèêî(çáåv)ëﬁî
' ∏ﬁÿØƒﬁçsçƒæØƒ & óÒà íuéÊìæ
    '================================================ 2009/07/22 Start
    'With VsGridK
    '   .Redraw = False
    '    If Ttl_Kensu + 1 < CisVsGridK.DispMax Then
    '        .Rows = CisVsGridK.DispMax
    '    Else
    '        .Rows = Ttl_Kensu + 1
    '    End If
    '
    '    ID = 0
    '    Do Until Not HKM_RDSTS
    '        ID = ID + 1
    '        'áÇ
    '        .TextMatrix(ID, 0) = Format(ID, "##")
    '        'çÏã∆
    '        .TextMatrix(ID, 1) = HKM.ã@äB
    '        If Trim(HKM.äÓñ{) <> "" Then
    '           'äÓñ{
    '           .TextMatrix(ID, 3) = "Åõ"
    '           'îıçl
    '           iB3_SJBiko = HKM.ã@äBîıçl
    '        End If
    '        Call HKMReadNext
    '    Loop
    '    Call HKMClose
    '
    '    For ID = 1 To .Rows - 1
    '        If Trim(.TextMatrix(ID, 1)) = "" Then Exit For
    '        If KMMMeisyoGet(.TextMatrix(ID, 1), 1) Then
    '            .TextMatrix(ID, 2) = KMM_Meisyo
    '        End If
    '    Next ID
    '
    '     .Row = 1: .Col = 1: .ColSel = .Cols - 1
    '     .Redraw = True
    'End With

    '================================================ 2009/07/22 End
End Sub
'+-------------------------------------------------------- 2004/12/13 Insert End
'+---------------------------------------+
'+  ÉOÉäÉbÉh (Click) éÊà¯êÊèÓïÒ
'+---------------------------------------+
Private Sub VSGrid1_Click()
    With vsGrid1
        
'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2006/12/19 Insert Start
'        gRtn = Body1Chk
        BD1Chk_Flg = 1
        gRtn = Body1Chk
        BD1Chk_Flg = 0
'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2006/12/19 Insert End
        If Not gRtn Then
            DoEvents
'           mBodyErrFld.SetFocus
           Exit Sub
        End If
        If ProcHB = "B1" Or ProcHB = "B2" Then
           GridBack = ProcHB
           Call HeadBodyClear("B2")
           Call HeadBodyClear("B3")
           If Trim(.TextMatrix(.Row, 1)) <> "" Then
              iB2_Torcd.Tag = .Row
              Call CisComboTehai.ComboSet
'''2007/12/11DEL              Call CisComboHNou.ComboSet
'''2007/12/11DEL              Call CisComboHJyu.ComboSet
              Call CisComboHMei.ComboSet
'''2007/12/11DEL              Call CisComboJNou.ComboSet
'''2007/12/11DEL              Call CisComboJJyu.ComboSet
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
              Call CisComboJsya.ComboSet
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END

              Call GridToText
'*--------------------------------------------------<< 2004/06/29 Insert Start >>
           Else
              B2Op_Hinmoku(1).Value = True     '>ïiñ⁄ÇÃèâä˙ílÇæØƒ
'*--------------------------------------------------<< 2004/06/29 Insert End >>
           End If
        End If
        If ProcHB = "B2" Then
'*--------------------------------------------------<< 2004/06/29 Insert Start >>
           B2_Area1.SetFocus                '>åüçıÇÃÃß›∏ºÆ›Ç™ñ≥ï\é¶Ç…Ç»ÇÈÇÃÇñhÇÆà◊
'*--------------------------------------------------<< 2004/06/29 Insert End >>
           iB2_Torcd.SetFocus
        End If
        If ProcHB = "B3" Then
           If SyoriKB2 = "ADD" Then
              Call GridToText(.Row)
           End If
           iB3_Seban.SetFocus
        End If
    End With
End Sub
'+---------------------------------------+
'+  ÉOÉäÉbÉh (DblClick) éÊà¯êÊèÓïÒ
'+---------------------------------------+
Private Sub vsGrid1_DblClick()
    
    With vsGrid1
        If ProcHB = "B1" Or ProcHB = "B2" Then
            If Trim(.TextMatrix(.Row, 1)) <> "" Then
                If SyoriKB2 <> "DEL" Then
                    RowColFlg = True
                    Call CisVsGrid.vsColor(å¬ï )
                    RowColFlg = False
                    
                    SyoriKB2 = "MNT"
                    Call SyoriName2(SyoriKB2)
                    ProcHB = "B3"
                    Call DispChange(ProcHB)
                    ' ∏ﬁÿØƒﬁì‡óeï\é¶å„ÅA√∑Ωƒì¸óÕ
                    'Call CisVsGridS.EditVisible("C")                                   2009/07/22
                    '+------------------------------------------- 2004/12/13 Insert
                    'Call CisVsGridK.EditVisible("C")                                   2009/07/22
                    '+------------------------------------------- 2004/12/13 Insert End
                    iB3_Seban.SetFocus
                End If
            Else
                If ProcHB = "B1" Then
                   Call PB_CHG_Click
                End If
            End If
        End If
    End With
End Sub
'+---------------------------------------+
'+  ÉOÉäÉbÉh (GotFocus) éÊà¯êÊèÓïÒ
'+---------------------------------------+
Private Sub vsGrid1_GotFocus()
    If ProcHB = "B2" Then
        iB2_Torcd.SetFocus
    End If
    If ProcHB = "B3" Then
        iB3_Seban.SetFocus
    End If
End Sub
'+---------------------------------------+
'+  ÉOÉäÉbÉh (GotFocus) éÊà¯êÊèÓïÒ
'+---------------------------------------+
Private Sub vsGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
       Call VSGrid1_Click
       If Not gRtn Then Exit Sub
       Call vsGrid1_DblClick
    End If
End Sub
'+---------------------------------------+
'+  ÉOÉäÉbÉh (RowColChange) éÊà¯êÊèÓïÒ
'+---------------------------------------+
Private Sub vsGrid1_RowColChange()
    If Not FormAct Then Exit Sub
    If RowColFlg Then Exit Sub
    Call VSGrid1_Click
End Sub
'=============================================== *** 2009/07/22 Start
''+==========================================================+
''+      ÉOÉäÉbÉhèàóù(çÏã∆èÓïÒ)                               +
''+==========================================================+
''+----------------------------------------------------------+
''+   å©èoÇµæØƒ & èâä˙èÛë‘æØƒ
''+----------------------------------------------------------+
'Private Sub GridInitS()
'    With CisVsGridS
''Å@∏ﬁÿØƒﬁèâä˙âªèàóù
''        .Clear
''       Call .FixedSet(å©èoÇµà íuê›íË,óÒà íuê›íË,œ∞ºﬁèàóù,
''                      æŸåÖêî,óÒî‘çÜ,å©èoÇµì‡óe,ï™óﬁ¿≤Ãﬂ)
'        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 2, 0, "No")
'        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 4, 1, "çÏã∆")
'        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 9, 2, "çÏ ã∆ ñº")
'        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 14, 3, "îıçl")
''Å@∏ﬁÿØƒﬁëIëèàóùópÅ@ëﬁî¥ÿ±
'        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 0, 4, "Select")
''  ï\é¶ ﬂ¿∞›æØƒ
''       Call .PatanSet( ﬂ¿∞›NO,ï\é¶ÉpÉ^Å[Éì)
'        Call .PatanSet(0, "0,1,2,3")
''  ï“èWì‡óeÉZÉbÉg
''       Call .EditSet(√∑ΩƒµÃﬁºﬁ™∏ƒ,∂◊—NO)
'        Call .EditSet(iB3_No, 0, 0, iB3_No.cFormat)
'        Call .EditSet(iB3_Sagyo, 1)
'        Call .EditSet(iB3_SBiko, 3)
'        .InitDisp
'    End With
'End Sub
'
'
''+----------------------+
''+  VSGridS_Click
''+----------------------+
'Private Sub VSGridS_Click()
'    Call CisVsGridS.EditEnable
'    Call CisVsGridS.EditVisible
'End Sub
''+---------------------------------------------------------+
''+          ÉOÉäÉbÉhï“èWÉGÉäÉA
''+---------------------------------------------------------+
''+----------------------+
''+  iB3_No
''+----------------------+
'Private Sub iB3_No_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call CisVsGridS.KeyEdit(iB3_No, KeyCode, Shift)
'End Sub
''+----------------------+
''+  iB3_Sagyo
''+----------------------+
'Private Sub iB3_Sagyo_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call CisVsGridS.KeyEdit(iB3_Sagyo, KeyCode, Shift)
'End Sub
''+-----------------------------+
''+  iB1_NKotei
''+-----------------------------+
'Private Sub iB3_Sagyo_GotFocus()
'    PB_Look.Tag = ActiveControl.Name
'    PB_Look.Visible = True
'End Sub
'Private Sub iB3_Sagyo_LostFocus()
'    If ActiveControl.Name <> "PB_Look" Then
'        PB_Look.Tag = ""
'        PB_Look.Visible = False
'    End If
'    With vsGridS
'        .TextMatrix(.Row, 2) = ""
'        If MeisyoGet("çÏã∆ä«óù", iB3_Sagyo) Then
'            .TextMatrix(.Row, 2) = SYM_Meisyo
'        End If
'    End With
'End Sub
''+----------------------+
''+  iB3_SBiko
''+----------------------+
'Private Sub iB3_SBiko_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call CisVsGridS.KeyEdit(iB3_SBiko, KeyCode, Shift)
'End Sub
''+----------------------+
''+  Dummy2
''+----------------------+
'Private Sub iB3_Biko_LostFocus()
'    Set ActObj = iB3_Biko
'End Sub
''+----------------------+
''+  Dummy2
''+----------------------+
'Private Sub Dummy2_GotFocus()
'    Call CisVsGridS.KeyEdit(iB3_No, vbKeyDown, 0)
'End Sub
''+----------------------+
''+  Dummy3
''+----------------------+
'Private Sub Dummy3_GotFocus()
'    If Not ActObj Is Nothing Then
'       Set ActObj = Nothing
'       iB3_No.SetFocus
'       Exit Sub
'    End If
'    If vsGridS.Row = 1 Then
'        '+------------------------------- 2004/12/13 Delete
''       If iB3_Biko.Enabled Then
''          iB3_Biko.SetFocus
''       End If
'        '+------------------------------- 2004/12/13 Delete End
'        '+------------------------------- 2004/12/13 Insert
'       If iB3_SJBiko.Enabled Then
'          iB3_SJBiko.SetFocus
'       End If
'        '+------------------------------- 2004/12/13 Insert End
'       Exit Sub
'    End If
'    Call CisVsGridS.KeyEdit(iB3_SBiko, vbKeyUp, 0)
'End Sub
''+----------------------+
''+  iB3_SJBiko
''+----------------------+
'Private Sub iB3_SJBiko_LostFocus()
'    Set ActObj = iB3_SJBiko
'End Sub
''+----------------------+
''+  VSGrid1_Scroll
''+----------------------+
'Private Sub vsGrid1_Scroll()
'    'ÉXÉNÉçÅ[ÉãéûÇÕÉeÉLÉXÉgÉ{ÉbÉNÉXÇîÒï\é¶Ç…ÇµÇ‹Ç∑ÅB
'    Call CisVsGrid.EditEnable
'End Sub
''+------------------------------------------------------------ 2004/12/13 Insert
''+==========================================================+
''+      ÉOÉäÉbÉhèàóù(ã@äBèÓïÒ)                               +
''+==========================================================+
''+==========================================================+
''+      ÉOÉäÉbÉhèàóù                                         +
''+==========================================================+
''+----------------------------------------------------------+
''+   å©èoÇµæØƒ & èâä˙èÛë‘æØƒ
''+----------------------------------------------------------+
'Private Sub GridInitK()
'    With CisVsGridK
''Å@∏ﬁÿØƒﬁèâä˙âªèàóù
''        .Clear
''       Call .FixedSet(å©èoÇµà íuê›íË,óÒà íuê›íË,œ∞ºﬁèàóù,
''                      æŸåÖêî,óÒî‘çÜ,å©èoÇµì‡óe,ï™óﬁ¿≤Ãﬂ)
'        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 0, 0, "No")
'        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 1, "ã@äBáÇ")
'        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 17, 2, "ã@äBñº")
'        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 3, "äÓñ{")
''Å@∏ﬁÿØƒﬁëIëèàóùópÅ@ëﬁî¥ÿ±
'        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 0, 4, "Select")
''  ï\é¶ ﬂ¿∞›æØƒ
''       Call .PatanSet( ﬂ¿∞›NO,ï\é¶ÉpÉ^Å[Éì)
'        Call .PatanSet(0, "0,1,2,3")
''  ï“èWì‡óeÉZÉbÉg
''       Call .EditSet(√∑ΩƒµÃﬁºﬁ™∏ƒ,∂◊—NO)
'        Call .EditSet(iB3_Kikai, 1)
'        .InitDisp
'    End With
'End Sub
''+----------------------+
''+  VSGridK_Click
''+----------------------+
'Private Sub VSGridK_Click()
'    Call CisVsGridK.EditEnable
'    Call CisVsGridK.EditVisible
'    Call iB3_Kikai_LostFocus
'End Sub
'Private Sub VsGridK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    With VsGridK
'        Call iB3_Kikai_LostFocus
'
'        If .MouseCol <> 3 Then Exit Sub
'        If Trim(.TextMatrix(.MouseRow, 1)) <> "" Then
'            For gLong = 1 To .Rows - 1
'                If gLong = .MouseRow And .MouseCol = 3 And Trim(.TextMatrix(.MouseRow, 1)) <> "" Then
'                   If Trim(.TextMatrix(.MouseRow, 3)) = "" Then
'                      .TextMatrix(.MouseRow, 3) = "Åõ"
'                   Else
'                      .TextMatrix(.MouseRow, 3) = ""
'                   End If
'                Else
'                  .TextMatrix(gLong, 3) = ""
'                End If
'            Next gLong
'        End If
'    End With
'End Sub
''+---------------------------------------------------------+
''+          ÉOÉäÉbÉhï“èWÉGÉäÉA
''+---------------------------------------------------------+
''+----------------------+
''+  iB3_Kikai
''+----------------------+
'Private Sub iB3_Kikai_KeyDown(KeyCode As Integer, Shift As Integer)
''    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
''       With VsGridK
''           .TextMatrix(.Row, 2) = ""
''           If KMMMeisyoGet(iB3_Kikai) Then
''               .TextMatrix(.Row, 2) = KMM_Meisyo
''           End If
''       End With
''    End If
'
'    Call CisVsGridK.KeyEdit(iB3_Kikai, KeyCode, Shift)
'End Sub
'Private Sub iB3_Kikai_GotFocus()
'    PB_Look.Tag = ActiveControl.Name
'    PB_Look.Visible = True
'End Sub
'Private Sub iB3_Kikai_LostFocus()
'    Set ActObj = iB3_Kikai
'    If ActiveControl.Name <> "PB_Look" Then
'        PB_Look.Tag = ""
'        PB_Look.Visible = False
'    End If
'    With VsGridK
'        .TextMatrix(.Row, 2) = ""
'        If KMMMeisyoGet(iB3_Kikai) Then
'            .TextMatrix(.Row, 2) = KMM_Meisyo
'        End If
'    End With
'End Sub
'''+----------------------+
'''+  DummyKS
'''+----------------------+
'Private Sub DummyKS_GotFocus()
'    If ActObj Is iB3_Biko Then
'       Set ActObj = Nothing
'       iB3_Kikai.SetFocus
'       Exit Sub
'    End If
'    If VsGridK.Row = 1 Then
'       If iB3_Biko.Enabled Then
'          iB3_Biko.SetFocus
'       End If
'       Exit Sub
'    End If
'    Call CisVsGridK.KeyEdit(iB3_Kikai, vbKeyUp, 0)
'End Sub
''+----------------------+
''+  DummyKE
''+----------------------+
'Private Sub DummyKE_GotFocus()
'    If ActObj Is iB3_Kikai Then
'       If Trim(iB3_Kikai) = "" Then
'          On Error Resume Next
'          If iB3_SJBiko.Enabled Then
'             iB3_SJBiko.SetFocus
'          End If
'          On Error GoTo 0
'          Exit Sub
'       End If
'       Call CisVsGridK.KeyEdit(iB3_Kikai, vbKeyDown, 0)
'    End If
'    If Not ActObj Is Nothing Then
'       Set ActObj = Nothing
'       iB3_Kikai.SetFocus
'       Exit Sub
'    End If
'End Sub
'''2007/12/11DEL+----------------------+
'''2007/12/11DEL+  VSGrid1_Scroll
'''2007/12/11DEL+----------------------+
''Private Sub vsGrid1_Scroll()
''    'ÉXÉNÉçÅ[ÉãéûÇÕÉeÉLÉXÉgÉ{ÉbÉNÉXÇîÒï\é¶Ç…ÇµÇ‹Ç∑ÅB
''    Call CisVsGrid.EditEnable
''End Sub
'+------------------------------------------------------------ 2004/12/13 Insert End
'=============================================== *** 2009/07/22 End

'======================================================== ***  2009/07/22 Start
'+==========================================================+
'+      ÉOÉäÉbÉhèàóù(çHíˆèÓïÒ)                              +
'+==========================================================+
'+==========================================================+
'+      ÉOÉäÉbÉhèàóù                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   å©èoÇµæØƒ & èâä˙èÛë‘æØƒ
'+----------------------------------------------------------+
Private Sub GridInit2()
    With CisVsGrid2
'Å@∏ﬁÿØƒﬁèâä˙âªèàóù
'        .Clear
'       Call .FixedSet(å©èoÇµà íuê›íË,óÒà íuê›íË,œ∞ºﬁèàóù,
'                      æŸåÖêî,óÒî‘çÜ,å©èoÇµì‡óe,ï™óﬁ¿≤Ãﬂ)
'       Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 0, 0, "No")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 4, 0, "çHèá")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 12, 1, "ã@äB(çÏã∆)")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 12, 2, "äOíçâ¡çH")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 12, 3, "ã@äBáÇ(édól)")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 6, 4, "€ŒﬁØƒ")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 12, 5, "çHíˆ")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 9, 6, "äÓèÄçHêî1")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 9, 7, "äÓèÄçHêî2")
'Å@∏ﬁÿØƒﬁëIëèàóùópÅ@ëﬁî¥ÿ±
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 0, 8, "Select")
'  ï\é¶ ﬂ¿∞›æØƒ
'       Call .PatanSet( ﬂ¿∞›NO,ï\é¶ÉpÉ^Å[Éì)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7")
        .InitDisp
    End With
End Sub
'***********************************************'
'*****       ÉOÉäÉbÉhì‡óeÉZÉbÉg(çHèáèÓïÒ)  *****
'***********************************************'
Private Sub GridSet2()
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
    Dim Kosu_Goukei As Currency
' íäèoèàóù èâä˙æØƒ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' íäèoåèêî(çáåv)ëﬁî
     
    CisVsGrid2.Clear

' ∏ﬁÿØƒﬁçsçƒæØƒ & óÒà íuéÊìæ
    With vsGrid2
       .Redraw = False
        If Ttl_Kensu + 1 < CisVsGrid2.DispMax Then
            .Rows = CisVsGrid2.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Kosu_Goukei = 0
    Do Until Not HKT_RDSTS
        ID = ID + 1
            
        Call HKTInfoToGrid2(ID)
        Kosu_Goukei = HKT.äÓèÄçHêî1
        
        Call HKTReadNext(1)
    Loop
    With vsGrid2
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    
    If Kosu_Goukei <> 0 Then
       B3lb_KosuKei = Format(Kosu_Goukei, "##0.000")
    Else
       B3lb_KosuKei = ""
    End If
End Sub
'+-------------------------------------+
'+  ïiî‘çHèáèÓïÒÇÊÇËÉOÉäÉbÉh2Ç÷ÉZÉbÉg    +
'+-------------------------------------+
Private Sub HKTInfoToGrid2(SetRow As Long)
    With vsGrid2
        For mIx = 0 To .Cols - 1
            .TextMatrix(SetRow, mIx) = ""
        Next mIx
        
        ' çHèá
        .TextMatrix(SetRow, 0) = Format(HKT.çHèá, "##0")
        ' ã@äB(çÏã∆)
        If RTrim(HKT.ã@äB) <> "" Then
           .TextMatrix(SetRow, SCol(1)) = RTrim(HKT.ã@äB) & ":" & RTrim(HKT.ã@äBñº)
        Else
           .TextMatrix(SetRow, SCol(1)) = ""
        End If
        ' ì‡äO/äOíçâ¡çH
        If HKT.äOíçâ¡çH = 1 Then
           .TextMatrix(SetRow, SCol(2)) = RTrim(HKT.äOíçâ¡çHêÊCD) & "ÅF" & RTrim(HKT.äOíçâ¡çHñº)
        Else
           .TextMatrix(SetRow, SCol(2)) = ""
        End If
        ' ã@äBNO(édól)
        If RTrim(HKT.ã@äBNO) <> "" Then
           .TextMatrix(SetRow, SCol(3)) = RTrim(HKT.ã@äBNO) & ":" & RTrim(HKT.ã@äBNOñº)
        Else
           .TextMatrix(SetRow, SCol(3)) = ""
        End If
        ' €ŒﬁØƒ/íPî≠
        If HKT.ÉçÉ{ÉbÉgãÊï™ = 0 Then
           .TextMatrix(SetRow, SCol(4)) = "€ŒﬁØƒ"
        Else
           .TextMatrix(SetRow, SCol(4)) = "íPî≠"
        End If
        ' çHíˆ
        If RTrim(HKT.çHíˆ) <> "" Then
           .TextMatrix(SetRow, SCol(5)) = RTrim(HKT.çHíˆ) & ":" & RTrim(HKT.çHíˆñº)
        Else
           .TextMatrix(SetRow, SCol(5)) = ""
        End If
        ' äÓèÄéûä‘1
        If RTrim(HKT.äÓèÄçHêîñº1) <> "" Then
           .TextMatrix(SetRow, SCol(6)) = RTrim(HKT.äÓèÄçHêîñº1) & ":" & Format(HKT.äÓèÄçHêî1, "#0.000")
        Else
           .TextMatrix(SetRow, SCol(6)) = Format(HKT.äÓèÄçHêî1, "#0.000")
        End If
        ' äÓèÄéûä‘2
        If RTrim(HKT.äÓèÄçHêîñº2) <> "" Then
           .TextMatrix(SetRow, SCol(7)) = RTrim(HKT.äÓèÄçHêîñº2) & ":" & Format(HKT.äÓèÄçHêî2, "#0.000")
        Else
           .TextMatrix(SetRow, SCol(7)) = Format(HKT.äÓèÄçHêî2, "#0.000")
        End If
    End With
End Sub
'======================================================== ***  2009/07/22 End

'+---------------------------------------+
'+      ñºèÃÉ}ÉXÉ^èâä˙ê›íË
'+---------------------------------------+
Private Sub MeisyoInit()
' ïiî‘ï“èW
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "ïiî‘ï“èW"
        .ãÊï™åÖêî = 1
        .íl = "1"
        .âpêîéö1 = "552"
        .ílñºèÃ = "ïiî‘ï“èWèâä˙ï\é¶"
        .çÏê¨é“ = gTanto
    End With
    If Not SYMInitCreate(True) Then
    End If
' çHíˆ
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "çHíˆä«óù"
        .ãÊï™åÖêî = 3
        .íl = "999"
        .ílñºèÃ = "ÇªÇÃëº"
        .çÏê¨é“ = gTanto
        .åüçıñº = "çH Å@íˆÅ@ åüÅ@ çı"
        .ÉRÅ[Éhñº = "çHíˆ"
        .ñºèÃñº = "çHÅ@íˆÅ@ñº"
    End With
    If Not SYMInitCreate(True) Then
    End If
' íSìñ
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "íSìñä«óù"
        .ãÊï™åÖêî = 3
        .íl = "999"
        .ílñºèÃ = "ÇªÇÃëº"
        .çÏê¨é“ = gTanto
        .åüçıñº = "íS Å@ìñÅ@ åüÅ@ çı"
        .ÉRÅ[Éhñº = "íSìñ"
        .ñºèÃñº = "íSÅ@ìñÅ@ñº"
    End With
    If Not SYMInitCreate(True) Then
    End If
' éËîzãÊï™
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "éËîzãÊï™"
        .ãÊï™åÖêî = 1
        .íl = "0"
        .ílñºèÃ = "çwì¸"
        .âpêîéö1 = "NJ "
        .êßå‰ãÊï™ = "X"
        .çÏê¨é“ = gTanto
        .åüçıñº = "éË îz ãÊ ï™ åü çı"
        .ÉRÅ[Éhñº = "éËîzãÊ"
        .ñºèÃñº = "éË îz ãÊ ï™ ñº"
        If Not SYMInitCreate(True) Then
            .íl = "1"
            .ílñºèÃ = "íºëó"
            .âpêîéö1 = "NJS"
            Call SYMInitCreate
            .íl = "2"
            .ílñºèÃ = "é–ì‡"
            .âpêîéö1 = "   "
            Call SYMInitCreate
            .íl = "3"
            .ílñºèÃ = "éÛãã"
            .âpêîéö1 = "NJS "
            Call SYMInitCreate
            '-----* 2008/06/09 Start
            .íl = "4"
            .ílñºèÃ = "éxãã"
            .âpêîéö1 = "N"
            Call SYMInitCreate
            '-----* 2008/06/09 End
        End If
    End With
' ç›å…ãÊï™
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "ç›å…ãÊï™"
        .ãÊï™åÖêî = 1
        .íl = "0"
        .êßå‰ãÊï™ = "X"
        .ílñºèÃ = "Ç»Çµ"
        .çÏê¨é“ = gTanto
        .åüçıñº = "ç› å… ãÊ ï™ åü çı"
        .ÉRÅ[Éhñº = "ç›å…ãÊ"
        .ñºèÃñº = "ç› å… ãÊ ï™ ñº"
        If Not SYMInitCreate(True) Then
            .íl = "1"
            .ílñºèÃ = "Ç†ÇË"
            Call SYMInitCreate
            .íl = "2"
            .ílñºèÃ = "åWêî"
            Call SYMInitCreate
            .íl = "3"
            .ílñºèÃ = "î≠íçì_"
            Call SYMInitCreate
            .íl = "4"
            .ílñºèÃ = "ì˙êî"
            Call SYMInitCreate
        End If
    End With
' çÏã∆
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "çÏã∆ä«óù"
        .ãÊï™åÖêî = 4
        .íl = "9999"
        .ílñºèÃ = "ÇªÇÃëº"
        .çÏê¨é“ = gTanto
        .åüçıñº = "çÏÅ@ã∆Å@åüÅ@çı"
        .ÉRÅ[Éhñº = "çÏã∆"
        .ñºèÃñº = "çÏÅ@Å@ã∆Å@Å@ñº"
        If Not SYMInitCreate(True) Then
        End If
    End With
' é˚óeäÌ
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "é˚óeäÌä«óù"
        .ãÊï™åÖêî = 10
        .íl = "XXXXXXXXXX"
        .ílñºèÃ = "ÇªÇÃëº"
        .çÏê¨é“ = gTanto
        .åüçıñº = "é˚ óe äÌ åü çı"
        .ÉRÅ[Éhñº = "é˚óeäÌ"
        .ñºèÃñº = "é˚  óe  äÌ  ñº"
        If Not SYMInitCreate(True) Then
        End If
    End With
' ïiî‘ãÊï™
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "ïiî‘ãÊï™"
        .ãÊï™åÖêî = 1
        .íl = "1"
        .êßå‰ãÊï™ = ""
        .ílñºèÃ = "êªïi"
        .çÏê¨é“ = gTanto
        .åüçıñº = "ïi î‘ ãÊ ï™ åü çı"
        .ÉRÅ[Éhñº = "ïiî‘ãÊ"
        .ñºèÃñº = "ïi î‘ ãÊ ï™ ñº"
        If Not SYMInitCreate(True) Then
            .íl = "2"
            .ílñºèÃ = "ïîïi"
            Call SYMInitCreate
        End If
    End With
End Sub

'-------------------------------------------------------' 2009.10.20 add start
'-------------------------------------------+
'+  éËîzãÊï™Ç™é–ì‡Å^ëºÇ≈ÇÃçÄñ⁄ï\é¶Å^îÒï\é¶  +
'-------------------------------------------+
Private Sub SyanaiTehai(kind As Boolean)
   'Set Image1.Picture = LoadPicture(gIniExe & "Image\Œﬂ≤›ƒ(AL).JPG")
   'Set Picture2.Picture = LoadPicture(gIniExe & "Image\Œﬂ≤›ƒ(AL).JPG")

'''''-------------------------------' 2009.10.23 add start
''''Back_HachuPHA.Visible = True    '
''''Call iB3_Point_LostFocus
''''Call iB3_HenCD_LostFocus
''''Call iB3_AtoCD_LostFocus
''''Exit Sub
'''''-------------------------------' 2009.10.23 add end
    
    If kind Then
''''        Back_Hachu(0).Top = Back_HachuAdrs.Top          '---------------' 2009.10.23 del start
''''        Back_Hachu(0).Left = Back_HachuAdrs.Left
''''        Back_Hachu(1).Top = Back_HachuAdrs.Top
''''        Back_Hachu(1).Left = Back_HachuAdrs.Left + Back_Hachu(0).Width
''''        Back_Hachu(1).Width = Back_HachuAdrs.Width - Back_Hachu(0).Width
''''        Back_HachuAdrs.Visible = False                  '---------------' 2009.10.23 del end
        Back_HachuPHA.Visible = True
        Call iB3_Point_LostFocus
        Call iB3_HenCD_LostFocus
        Call iB3_AtoCD_LostFocus
    Else
''''        Back_Hachu(0).Left = PosiSave0X                 '---------------' 2009.10.23 del start
''''        Back_Hachu(0).Top = PosiSave0Y
''''        Back_Hachu(1).Left = PosiSave1X
''''        Back_Hachu(1).Top = PosiSave1Y
''''        Back_Hachu(1).Width = PosiSave1W
''''        Back_HachuAdrs.Visible = True                   '---------------' 2009.10.23 del end
        Back_HachuPHA.Visible = False
    End If
End Sub
'-------------------------------------------------------' 2009.10.20 add end

