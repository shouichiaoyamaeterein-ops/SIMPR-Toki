VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXM0070 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "çﬁóøÉ}ÉXÉ^ÉÅÉìÉeÉiÉìÉX"
   ClientHeight    =   11175
   ClientLeft      =   1725
   ClientTop       =   1740
   ClientWidth     =   15300
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
   ScaleHeight     =   11175
   ScaleWidth      =   15300
   WindowState     =   2  'ç≈ëÂâª
   Begin VB.Timer Timer4 
      Left            =   150
      Top             =   -1050
   End
   Begin VB.PictureBox B1_Area1B 
      Height          =   5295
      Left            =   240
      ScaleHeight     =   5235
      ScaleWidth      =   10995
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   5820
      Width           =   11055
      Begin VB.PictureBox B3_Area3 
         Height          =   675
         Left            =   7230
         ScaleHeight     =   615
         ScaleWidth      =   2205
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   -720
         Width           =   2265
         Begin Cis3D_v60.CIS3D PB_KTN 
            Height          =   615
            Left            =   0
            Top             =   15
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1085
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
            cPositionY      =   55
         End
         Begin Cis3D_v60.CIS3D PB_STN 
            Height          =   615
            Left            =   1110
            Top             =   15
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1085
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
            cPositionY      =   55
         End
      End
      Begin VB.PictureBox B2_Area1 
         BackColor       =   &H00C0C0C0&
         Height          =   645
         Left            =   60
         ScaleHeight     =   585
         ScaleWidth      =   5055
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   30
         Width           =   5115
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   465
            Left            =   60
            Top             =   60
            Width           =   3900
            _ExtentX        =   6879
            _ExtentY        =   820
            BackColor       =   12632319
            ForeColor       =   16711680
            Caption         =   "édì¸êÊ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Begin CisText_V60.CisText iB2_Torcd 
               Height          =   330
               Left            =   810
               TabIndex        =   23
               Top             =   60
               Width           =   990
               _ExtentX        =   1746
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
               Text            =   "XXXXXXX"
               MaxLength       =   7
            End
            Begin Cis3D_v60.CIS3D B2lb_Torcd 
               Height          =   330
               Left            =   1830
               Top             =   60
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   582
               BackColor       =   16777152
               Caption         =   "XXXXXXXXX1XXXX"
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
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   30
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   465
            Index           =   9
            Left            =   3960
            Top             =   60
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   820
            BackColor       =   12632319
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
            cPositionX      =   30
            Begin CisText_V60.CisText iB2_Ukeir 
               Height          =   330
               Left            =   510
               TabIndex        =   25
               Top             =   60
               Width           =   405
               _ExtentX        =   714
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
               Text            =   "XX"
               MaxLength       =   2
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   465
            Index           =   11
            Left            =   4980
            Top             =   -480
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   820
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
            cPositionX      =   30
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
               Height          =   315
               Index           =   0
               Left            =   540
               TabIndex        =   104
               Top             =   90
               Width           =   735
            End
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
               Height          =   315
               Index           =   1
               Left            =   1290
               TabIndex        =   103
               Top             =   90
               Value           =   -1  'True
               Width           =   720
            End
         End
      End
      Begin Cis3D_v60.CIS3D K_Sykbnm2 
         Height          =   615
         Left            =   9510
         Top             =   30
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1085
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
         cBoderWidth     =   4
         cAlingnment     =   7
      End
      Begin VB.PictureBox B3_Area1 
         BackColor       =   &H00C0C0C0&
         Height          =   4560
         Left            =   60
         ScaleHeight     =   4500
         ScaleWidth      =   10815
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   690
         Width           =   10875
         Begin Cis3D_v60.CIS3D CIS3D12 
            Height          =   735
            Left            =   4200
            Top             =   825
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   1296
            ForeColor       =   16711680
            Caption         =   "íIâµä«óù"
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
            Begin VB.CheckBox B3Chk_Tana 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Ç∑ÇÈ"
               Height          =   240
               Left            =   240
               TabIndex        =   43
               Top             =   405
               Width           =   1065
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   735
            Index           =   5
            Left            =   8880
            Top             =   825
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   1296
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
               TabIndex        =   48
               Top             =   330
               Width           =   1770
               _ExtentX        =   3122
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
            Height          =   735
            Index           =   34
            Left            =   9000
            Top             =   3720
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   1296
            ForeColor       =   16711680
            Caption         =   "çﬁéøï\é¶"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Begin VB.CheckBox B3Chk_JIS 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ïWèÄ"
               BeginProperty Font 
                  Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   300
               TabIndex        =   82
               Top             =   360
               Width           =   1230
            End
            Begin VB.CheckBox B3Chk_GenPn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ëÂ"
               BeginProperty Font 
                  Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   405
               TabIndex        =   84
               Top             =   360
               Width           =   540
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   735
            Index           =   33
            Left            =   9000
            Top             =   3720
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   1296
            ForeColor       =   16711680
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionX      =   30
            cPositionY      =   30
            Begin VB.CheckBox B3Chk_KanBn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ëÂ"
               BeginProperty Font 
                  Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   210
               TabIndex        =   83
               Top             =   360
               Width           =   600
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   735
            Index           =   42
            Left            =   7665
            Top             =   1560
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   1296
            ForeColor       =   16711680
            Caption         =   "ìæÅ@à”Å@êÊ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Begin CisText_V60.CisText iB3_TTorcd 
               Height          =   360
               Left            =   120
               TabIndex        =   53
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
            Begin Cis3D_v60.CIS3D B3lb_TTorcd 
               Height          =   345
               Left            =   1110
               Top             =   330
               Width           =   1800
               _ExtentX        =   3175
               _ExtentY        =   609
               BackColor       =   16777152
               Caption         =   "XXXXXXXXX1XXXX"
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
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   30
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   735
            Index           =   41
            Left            =   4335
            Top             =   1560
            Width           =   3330
            _ExtentX        =   5874
            _ExtentY        =   1296
            ForeColor       =   16711680
            Caption         =   "â¡Å@çHÅ@êÊ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cPositionX      =   180
            cPositionY      =   30
            Begin CisText_V60.CisText iB3_STorcd 
               Height          =   345
               Left            =   120
               TabIndex        =   52
               Top             =   330
               Width           =   990
               _ExtentX        =   1746
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
               Text            =   "XXXXXXX"
               MaxLength       =   7
            End
            Begin Cis3D_v60.CIS3D B3lb_STorcd 
               Height          =   345
               Left            =   1110
               Top             =   330
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   609
               BackColor       =   16777152
               Caption         =   "XXXXXXXXX1XXXX"
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
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   30
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   735
            Index           =   15
            Left            =   1440
            Top             =   3720
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   1296
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
               Left            =   105
               TabIndex        =   76
               Top             =   390
               Width           =   1035
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   630
            Index           =   32
            Left            =   7200
            Top             =   2295
            Width           =   3570
            _ExtentX        =   6297
            _ExtentY        =   1111
            ForeColor       =   16711680
            Caption         =   "é©Å@é–Å@ïiÅ@î‘"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Begin CisText_V60.CisText iB3_SHinbn 
               Height          =   300
               Left            =   150
               TabIndex        =   56
               Top             =   255
               Width           =   3300
               _ExtentX        =   5821
               _ExtentY        =   529
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
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   630
            Index           =   29
            Left            =   60
            Top             =   2295
            Width           =   3570
            _ExtentX        =   6297
            _ExtentY        =   1111
            ForeColor       =   16711680
            Caption         =   "î≠Å@íçÅ@ïiÅ@î‘"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            cPositionX      =   100
            Begin CisText_V60.CisText iB3_HHinbn 
               Height          =   300
               Left            =   150
               TabIndex        =   54
               Top             =   255
               Width           =   3300
               _ExtentX        =   5821
               _ExtentY        =   529
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
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   630
            Index           =   30
            Left            =   3630
            Top             =   2295
            Width           =   3570
            _ExtentX        =   6297
            _ExtentY        =   1111
            ForeColor       =   16711680
            Caption         =   "éÛÅ@ããÅ@ïiÅ@î‘"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Begin CisText_V60.CisText iB3_JHinbn 
               Height          =   300
               Left            =   150
               TabIndex        =   55
               Top             =   255
               Width           =   3300
               _ExtentX        =   5821
               _ExtentY        =   529
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
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   420
            Index           =   31
            Left            =   60
            Top             =   5085
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   741
            ForeColor       =   16711680
            Caption         =   "íPâøä∑éZ"
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
            cPositionX      =   100
            Begin Cis3D_v60.CIS3D CIS3D24 
               Height          =   270
               Left            =   2595
               Top             =   120
               Width           =   3510
               _ExtentX        =   6191
               _ExtentY        =   476
               ForeColor       =   255
               Caption         =   "Å¶0/ñ¢ì¸óÕéûÇÕÇPÇ∆ÇµÇƒìoò^ÇµÇ‹Ç∑"
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cBoderStyle     =   2
            End
            Begin CisText_V60.CisText iB3_TKansn 
               Height          =   360
               Left            =   1095
               TabIndex        =   85
               Top             =   30
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   635
               cFormat         =   "###,##0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###,##0.000"
               MaxLength       =   11
               cDataType       =   2
               cDataReplace    =   1
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   735
            Index           =   25
            Left            =   5610
            Top             =   1560
            Width           =   5160
            _ExtentX        =   9102
            _ExtentY        =   1296
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
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   735
            Index           =   2
            Left            =   1260
            Top             =   1560
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   1296
            ForeColor       =   16711680
            Caption         =   "ÉÅÅ[ÉJÅ["
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cPositionX      =   180
            cPositionY      =   30
            Begin CisText_V60.CisText iB3_MTorcd 
               Height          =   360
               Left            =   30
               TabIndex        =   50
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
            Begin Cis3D_v60.CIS3D B3lb_MTorcd 
               Height          =   345
               Left            =   1005
               Top             =   330
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   609
               BackColor       =   16777152
               Caption         =   "XXXXXXXXX1XXXX"
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
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   30
            End
            Begin CisText_V60.CisText iB3_MUkeir 
               Height          =   360
               Left            =   2535
               TabIndex        =   51
               Top             =   330
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
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   465
               Index           =   34
               Left            =   2445
               Top             =   30
               Width           =   495
               _ExtentX        =   873
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
            Height          =   735
            Index           =   24
            Left            =   4335
            Top             =   1560
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   1296
            ForeColor       =   16711680
            Caption         =   "ì`ï[î≠çs"
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
               Left            =   240
               TabIndex        =   90
               Top             =   405
               Width           =   945
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   375
            Index           =   22
            Left            =   3885
            Top             =   5115
            Width           =   6885
            _ExtentX        =   12144
            _ExtentY        =   661
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
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   375
            Index           =   21
            Left            =   60
            Top             =   5115
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   661
            ForeColor       =   16711680
            Caption         =   "[ëOçHíˆî≠çs]î≠íçï\é¶"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            cPositionX      =   100
            Begin VB.CheckBox B3Chk_Maehak 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ï\é¶ñ≥"
               Height          =   300
               Left            =   2610
               TabIndex        =   86
               Top             =   60
               Width           =   1050
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   735
            Index           =   0
            Left            =   5640
            Top             =   825
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   1296
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
            Begin Cis3D_v60.CIS3D CIS3D16 
               Height          =   285
               Left            =   1530
               Top             =   360
               Width           =   195
               _ExtentX        =   344
               _ExtentY        =   503
               ForeColor       =   16711680
               Caption         =   "Å`"
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
               cAlingnment     =   6
               cBoderStyle     =   2
            End
            Begin CisYMD_With_Btn.CisYMDwB iB3_SDate 
               Height          =   330
               Left            =   60
               TabIndex        =   46
               Top             =   360
               Width           =   1455
               _ExtentX        =   2249
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
               Object.Width           =   1455
               cSize           =   1
               cChkResult      =   0   'False
               cCalenderPosition=   1
               cType           =   1
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
            Begin CisYMD_With_Btn.CisYMDwB iB3_EDate 
               Height          =   330
               Left            =   1710
               TabIndex        =   47
               Top             =   330
               Width           =   1455
               _ExtentX        =   2249
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
               Object.Width           =   1455
               cSize           =   1
               cChkResult      =   0   'False
               cCalenderPosition=   1
               cType           =   1
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
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   795
            Index           =   3
            Left            =   60
            Top             =   2925
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   1402
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
            Begin CisText_V60.CisText iB3_HRitu 
               Height          =   360
               Left            =   45
               TabIndex        =   57
               Top             =   330
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
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   375
            Index           =   12
            Left            =   7455
            Top             =   2925
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   661
            ForeColor       =   16711680
            Caption         =   "ójì˙"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            cPositionX      =   120
            Begin VB.CheckBox B3Chk_Week 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ã‡"
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
               Index           =   4
               Left            =   2670
               TabIndex        =   69
               Top             =   60
               Width           =   495
            End
            Begin VB.CheckBox B3Chk_Week 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ñÿ"
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
               Index           =   3
               Left            =   2175
               TabIndex        =   68
               Top             =   60
               Width           =   495
            End
            Begin VB.CheckBox B3Chk_Week 
               BackColor       =   &H00C0C0C0&
               Caption         =   "êÖ"
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
               Index           =   2
               Left            =   1680
               TabIndex        =   67
               Top             =   60
               Width           =   495
            End
            Begin VB.CheckBox B3Chk_Week 
               BackColor       =   &H00C0C0C0&
               Caption         =   "âŒ"
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
               Index           =   1
               Left            =   1185
               TabIndex        =   66
               Top             =   60
               Width           =   495
            End
            Begin VB.CheckBox B3Chk_Week 
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
               Height          =   285
               Index           =   0
               Left            =   720
               TabIndex        =   65
               Top             =   60
               Width           =   495
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   795
            Index           =   11
            Left            =   5865
            Top             =   2925
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   1402
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
            Begin CisText_V60.CisText iB3_Lot2 
               Height          =   360
               Left            =   60
               TabIndex        =   64
               Top             =   390
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   635
               cFormat         =   "###,##0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###,##0.000"
               MaxLength       =   11
               cDataType       =   2
               cDataReplace    =   1
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   795
            Index           =   9
            Left            =   3795
            Top             =   2925
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   1402
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
               Left            =   60
               TabIndex        =   62
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
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   795
            Index           =   8
            Left            =   3075
            Top             =   2925
            Width           =   720
            _ExtentX        =   1270
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   30
            Begin CisText_V60.CisText iB3_Keisu3 
               Height          =   360
               Left            =   60
               TabIndex        =   61
               Top             =   390
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
            Height          =   795
            Index           =   7
            Left            =   2355
            Top             =   2925
            Width           =   720
            _ExtentX        =   1270
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   30
            Begin CisText_V60.CisText iB3_Keisu2 
               Height          =   360
               Left            =   60
               TabIndex        =   60
               Top             =   390
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
            Height          =   795
            Index           =   6
            Left            =   1635
            Top             =   2925
            Width           =   720
            _ExtentX        =   1270
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   30
            Begin CisText_V60.CisText iB3_Keisu1 
               Height          =   360
               Left            =   60
               TabIndex        =   59
               Top             =   390
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
            Height          =   420
            Index           =   13
            Left            =   7455
            Top             =   3300
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   741
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
            cAlingnment     =   1
            cPositionX      =   240
            Begin CisText_V60.CisText iB3_Day 
               Height          =   360
               Index           =   0
               Left            =   690
               TabIndex        =   70
               Top             =   30
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
               Left            =   1140
               TabIndex        =   71
               Top             =   30
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
               Left            =   1590
               TabIndex        =   72
               Top             =   30
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
               Left            =   2040
               TabIndex        =   73
               Top             =   30
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
               Left            =   2490
               TabIndex        =   74
               Top             =   30
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
            Height          =   735
            Index           =   1
            Left            =   60
            Top             =   1560
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   1296
            BackColor       =   12632319
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
            Begin VB.ComboBox B3Comb_THKbn 
               BackColor       =   &H00C0E0FF&
               Height          =   360
               Left            =   90
               Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
               TabIndex        =   49
               Top             =   330
               Width           =   900
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D8 
            Height          =   795
            Left            =   3210
            Top             =   30
            Width           =   1125
            _ExtentX        =   1984
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
            cPositionY      =   20
            Begin CisText_V60.CisText iB3_Cycle1 
               Height          =   330
               Left            =   60
               TabIndex        =   29
               Top             =   390
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   582
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
               Height          =   330
               Left            =   300
               TabIndex        =   30
               Top             =   390
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   582
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
               Height          =   330
               Left            =   690
               TabIndex        =   31
               Top             =   390
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   582
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
            Height          =   795
            Left            =   8295
            Top             =   30
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   1402
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cPositionX      =   60
            cPositionY      =   60
            Begin CisText_V60.CisText iB3_Syuyo2 
               Height          =   330
               Left            =   90
               TabIndex        =   36
               Top             =   60
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   582
               cFormat         =   "###,##0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###,##0.000"
               MaxLength       =   11
               cDataType       =   2
               cDataReplace    =   1
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin CisText_V60.CisText iB3_Hako2 
               Height          =   330
               Left            =   360
               TabIndex        =   37
               Top             =   420
               Width           =   1335
               _ExtentX        =   2355
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
               Text            =   "XXXXXXXXXX"
               MaxLength       =   10
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   315
               Index           =   7
               Left            =   30
               Top             =   420
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   556
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
               cBoderColor1    =   12632256
               cBoderColor2    =   12632256
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cBoderStyle     =   2
               cPositionY      =   30
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D20 
            Height          =   795
            Left            =   6390
            Top             =   30
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   " ïœ çX ì˙ - ï÷"
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
            Begin CisYMD_With_Btn.CisYMDwB iB3_CDate 
               Height          =   330
               Left            =   30
               TabIndex        =   34
               Top             =   390
               Width           =   1455
               _ExtentX        =   2249
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
               Object.Width           =   1455
               cSize           =   1
               cChkResult      =   0   'False
               cCalenderPosition=   1
               cType           =   1
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
            Begin CisText_V60.CisText iB3_CBin 
               Height          =   330
               Left            =   1500
               TabIndex        =   35
               Top             =   390
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   582
               cFormat         =   "00"
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
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "##"
               cILength        =   2
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   795
            Left            =   4335
            Top             =   30
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1402
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cPositionX      =   900
            cPositionY      =   60
            Begin CisText_V60.CisText iB3_Syuyo1 
               Height          =   330
               Left            =   420
               TabIndex        =   32
               Top             =   60
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   582
               cFormat         =   "###,##0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###,##0.000"
               MaxLength       =   11
               cDataType       =   2
               cDataReplace    =   1
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
            Begin CisText_V60.CisText iB3_Hako1 
               Height          =   330
               Left            =   690
               TabIndex        =   33
               Top             =   420
               Width           =   1335
               _ExtentX        =   2355
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
               Text            =   "XXXXXXXXXX"
               MaxLength       =   10
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   315
               Index           =   16
               Left            =   360
               Top             =   420
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   556
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
               cBoderColor1    =   12632256
               cBoderColor2    =   12632256
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cBoderStyle     =   2
               cPositionY      =   30
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   735
               Index           =   15
               Left            =   90
               Top             =   30
               Width           =   270
               _ExtentX        =   476
               _ExtentY        =   1296
               ForeColor       =   16711680
               Caption         =   "é˚óeêî"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   11.25
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
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   735
            Index           =   18
            Left            =   60
            Top             =   825
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   1296
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
            Begin CisText_V60.CisText iB3_Okiba1 
               Height          =   375
               Left            =   105
               TabIndex        =   38
               Top             =   315
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XX"
               MaxLength       =   2
            End
            Begin CisText_V60.CisText iB3_Okiba2 
               Height          =   375
               Left            =   435
               TabIndex        =   39
               Top             =   315
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XX"
               MaxLength       =   2
            End
            Begin CisText_V60.CisText iB3_Okiba3 
               Height          =   375
               Left            =   765
               TabIndex        =   40
               Top             =   315
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XX"
               MaxLength       =   2
            End
            Begin CisText_V60.CisText iB3_Okiba4 
               Height          =   375
               Left            =   1095
               TabIndex        =   41
               Top             =   315
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   11.25
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
         Begin Cis3D_v60.CIS3D Back_Seban 
            Height          =   795
            Left            =   60
            Top             =   30
            Width           =   1200
            _ExtentX        =   2117
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
            Begin CisText_V60.CisText iB3_Seban 
               Height          =   360
               Left            =   60
               TabIndex        =   26
               Top             =   390
               Width           =   1095
               _ExtentX        =   1931
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
            Height          =   795
            Index           =   17
            Left            =   1260
            Top             =   30
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   1402
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
            Begin VB.OptionButton B3Op_SKKbn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "éwé¶"
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   1200
               TabIndex        =   28
               Top             =   450
               Width           =   720
            End
            Begin VB.OptionButton B3Op_SKKbn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Ç©ÇÒÇŒÇÒ"
               BeginProperty Font 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   27
               Top             =   450
               Value           =   -1  'True
               Width           =   1155
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D11 
            Height          =   735
            Left            =   1590
            Top             =   825
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   1296
            BackColor       =   12632319
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
               TabIndex        =   42
               Top             =   330
               Width           =   930
               _ExtentX        =   1640
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
               Height          =   345
               Left            =   1020
               Top             =   330
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   609
               BackColor       =   16777152
               Caption         =   "XXXXXXXXX1XXXX"
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
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   30
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   795
            Index           =   10
            Left            =   4275
            Top             =   2925
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   1402
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
            Begin CisText_V60.CisText iB3_Lot1 
               Height          =   360
               Left            =   60
               TabIndex        =   63
               Top             =   390
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   635
               cFormat         =   "###,##0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###,##0.000"
               MaxLength       =   11
               cDataType       =   2
               cDataReplace    =   1
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#####0.000"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   795
            Index           =   4
            Left            =   645
            Top             =   2925
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   1402
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
            cPositionX      =   30
            cPositionY      =   30
            Begin VB.CheckBox B3Chk_NHKbn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ÇµÇ»Ç¢"
               BeginProperty Font 
                  Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   45
               TabIndex        =   58
               Top             =   360
               Width           =   930
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D 
            Height          =   795
            Index           =   0
            Left            =   10020
            Top             =   30
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "ç≈èIáÇ"
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
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D B3lb_EndNO 
               Height          =   360
               Left            =   150
               Top             =   360
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   635
               BackColor       =   16777152
               Caption         =   "999"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
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
            Height          =   735
            Index           =   2
            Left            =   3180
            Top             =   825
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   1296
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
               Height          =   420
               Left            =   960
               TabIndex        =   45
               Top             =   300
               Visible         =   0   'False
               Width           =   450
            End
            Begin CisText_V60.CisText iB3_ADD 
               Height          =   360
               Left            =   60
               TabIndex        =   44
               Top             =   330
               Visible         =   0   'False
               Width           =   900
               _ExtentX        =   1588
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
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   735
            Index           =   14
            Left            =   60
            Top             =   3720
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   1296
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
            Begin CisText_V60.CisText iB3_Densy 
               Height          =   360
               Left            =   60
               TabIndex        =   75
               Top             =   330
               Width           =   300
               _ExtentX        =   529
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
            End
            Begin Cis3D_v60.CIS3D B3lb_Densy 
               Height          =   345
               Left            =   360
               Top             =   330
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   609
               BackColor       =   16777152
               Caption         =   "XXXXXXXX"
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
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   30
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   735
            Index           =   18
            Left            =   5310
            Top             =   3720
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   1296
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
            Begin VB.CheckBox B3Chk_JRKbn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "î≠íçéû"
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
               Left            =   180
               TabIndex        =   79
               Top             =   390
               Width           =   1095
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   735
            Index           =   19
            Left            =   6600
            Top             =   3720
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   1296
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
            Begin VB.ComboBox B3Comb_Meisi 
               BackColor       =   &H00C0E0FF&
               Height          =   360
               Left            =   60
               Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
               TabIndex        =   80
               Top             =   330
               Width           =   1800
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   735
            Index           =   20
            Left            =   8520
            Top             =   3720
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   1296
            ForeColor       =   16711680
            Caption         =   "PáÇ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Left            =   45
               TabIndex        =   81
               Top             =   330
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
            Height          =   735
            Index           =   17
            Left            =   4020
            Top             =   3720
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   1296
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
               Left            =   150
               TabIndex        =   78
               Top             =   390
               Width           =   1050
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   735
            Index           =   16
            Left            =   2730
            Top             =   3720
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   1296
            ForeColor       =   16711680
            Caption         =   "éxããèëî≠çs"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Left            =   135
               TabIndex        =   77
               Top             =   390
               Width           =   1050
            End
         End
         Begin Cis3D_v60.CIS3D Back_Juchu 
            Height          =   735
            Index           =   0
            Left            =   7410
            Top             =   825
            Visible         =   0   'False
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1296
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
            Begin CisText_V60.CisText iB3_BikoJ 
               Height          =   360
               Left            =   120
               TabIndex        =   106
               Top             =   330
               Width           =   3165
               _ExtentX        =   5583
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
            Index           =   2
            Left            =   60
            Top             =   1530
            Visible         =   0   'False
            Width           =   1980
            _ExtentX        =   3493
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
            Begin CisText_V60.CisText iB3_DensyJ 
               Height          =   360
               Left            =   120
               TabIndex        =   107
               Top             =   330
               Width           =   300
               _ExtentX        =   529
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
            End
            Begin Cis3D_v60.CIS3D B3lb_DensyJ 
               Height          =   360
               Left            =   420
               Top             =   330
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   635
               BackColor       =   16777152
               Caption         =   "XXXXXXXXX1XXXX"
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
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   30
            End
         End
         Begin Cis3D_v60.CIS3D Back_Juchu 
            Height          =   765
            Index           =   22
            Left            =   6315
            Top             =   1410
            Visible         =   0   'False
            Width           =   1470
            _ExtentX        =   2593
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
            Begin VB.CheckBox B3Chk_JRKbnJ 
               BackColor       =   &H00C0C0C0&
               Caption         =   "î≠íçéû"
               Height          =   285
               Left            =   180
               TabIndex        =   111
               Top             =   390
               Width           =   1095
            End
         End
         Begin Cis3D_v60.CIS3D Back_Juchu 
            Height          =   765
            Index           =   23
            Left            =   7785
            Top             =   1530
            Visible         =   0   'False
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   1349
            ForeColor       =   16711680
            Caption         =   "ÇoáÇ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Begin CisText_V60.CisText iB3_PNoJ 
               Height          =   360
               Left            =   195
               TabIndex        =   112
               Top             =   330
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
         Begin Cis3D_v60.CIS3D Back_Juchu 
            Height          =   765
            Index           =   24
            Left            =   4890
            Top             =   1410
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
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
            Begin VB.CheckBox B3Chk_PrtkbJJ 
               BackColor       =   &H00C0C0C0&
               Caption         =   "î≠çsóL"
               Height          =   300
               Left            =   150
               TabIndex        =   110
               Top             =   405
               Width           =   1035
            End
         End
         Begin Cis3D_v60.CIS3D Back_Juchu 
            Height          =   765
            Index           =   25
            Left            =   2040
            Top             =   1530
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   30
            Begin VB.CheckBox B3Chk_PrtkbNJ 
               BackColor       =   &H00C0C0C0&
               Caption         =   "î≠çsóL"
               Height          =   300
               Left            =   105
               TabIndex        =   108
               Top             =   405
               Width           =   1035
            End
         End
         Begin Cis3D_v60.CIS3D Back_Juchu 
            Height          =   765
            Index           =   26
            Left            =   3465
            Top             =   1530
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   1349
            ForeColor       =   16711680
            Caption         =   "éxããèëî≠çs"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   300
               Left            =   135
               TabIndex        =   109
               Top             =   405
               Width           =   1050
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   735
            Index           =   26
            Left            =   60
            Top             =   3720
            Width           =   9960
            _ExtentX        =   17568
            _ExtentY        =   1296
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
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   420
            Index           =   27
            Left            =   6195
            Top             =   5085
            Width           =   4575
            _ExtentX        =   8070
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
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   615
            Index           =   28
            Left            =   3630
            Top             =   2295
            Width           =   3570
            _ExtentX        =   6297
            _ExtentY        =   1085
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
         Begin Cis3D_v60.CIS3D Back_Juchu 
            Height          =   615
            Index           =   3
            Left            =   60
            Top             =   2295
            Width           =   10710
            _ExtentX        =   18891
            _ExtentY        =   1085
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
               Height          =   375
               Index           =   23
               Left            =   0
               Top             =   1530
               Width           =   10710
               _ExtentX        =   18891
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D CIS3D14 
                  Height          =   330
                  Left            =   525
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   582
                  ForeColor       =   16711680
                  Caption         =   "éÛããïiî‘"
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
               Begin Cis3D_v60.CIS3D CIS3D23 
                  Height          =   330
                  Left            =   5835
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   582
                  ForeColor       =   16711680
                  Caption         =   "î≠íçïiî‘"
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
            End
         End
         Begin Cis3D_v60.CIS3D Back_Juchu 
            Height          =   765
            Index           =   1
            Left            =   8535
            Top             =   1530
            Visible         =   0   'False
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   1349
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
         Begin Cis3D_v60.CIS3D Back_SHin2 
            Height          =   735
            Left            =   9570
            Top             =   3720
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   1296
            ForeColor       =   16711680
            Caption         =   "égópïiî‘2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Begin CisText_V60.CisText iB3_SHin2 
               Height          =   360
               Left            =   60
               TabIndex        =   113
               Top             =   390
               Width           =   1095
               _ExtentX        =   1931
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
         Begin Cis3D_v60.CIS3D Back_SHin1 
            Height          =   735
            Left            =   8370
            Top             =   3720
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   1296
            ForeColor       =   16711680
            Caption         =   "égópïiî‘1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Begin CisText_V60.CisText iB3_SHin1 
               Height          =   360
               Left            =   60
               TabIndex        =   114
               Top             =   390
               Width           =   1095
               _ExtentX        =   1931
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
      End
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   1680
      Left            =   4320
      ScaleHeight     =   1620
      ScaleWidth      =   10680
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   1650
      Width           =   10740
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   780
         Left            =   4980
         Top             =   810
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1376
         BackColor       =   12632319
         ForeColor       =   16711680
         Caption         =   "éËîzíPà "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_TTani 
            Height          =   360
            Left            =   120
            TabIndex        =   13
            Top             =   360
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
            cDataReplace    =   1
         End
         Begin Cis3D_v60.CIS3D B1lb_TTani 
            Height          =   345
            Left            =   480
            Top             =   360
            Width           =   735
            _ExtentX        =   1296
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   780
         Index           =   0
         Left            =   3315
         Top             =   810
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "ä∑éZíl"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_Kansn 
            Height          =   360
            Left            =   75
            TabIndex        =   12
            Top             =   360
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   635
            cFormat         =   "###,##0.000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "###,##0.000"
            MaxLength       =   11
            cDataType       =   2
            cDataReplace    =   1
            cFcDecimal      =   3
            cFdAutoFormat   =   1
            cGFormat        =   "#####0.000"
            cILength        =   6
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   780
         Left            =   2070
         Top             =   810
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "égópíPà "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_Tani 
            Height          =   360
            Left            =   105
            TabIndex        =   11
            Top             =   360
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
            cDataReplace    =   1
         End
         Begin Cis3D_v60.CIS3D B1lb_Tani 
            Height          =   345
            Left            =   465
            Top             =   360
            Width           =   690
            _ExtentX        =   1217
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   750
         Index           =   12
         Left            =   60
         Top             =   60
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   1323
         ForeColor       =   16711680
         Caption         =   "çﬁ óø ñº"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_Name 
            Height          =   360
            Left            =   90
            TabIndex        =   6
            Top             =   330
            Width           =   5085
            _ExtentX        =   8969
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
         Height          =   795
         Index           =   1
         Left            =   12330
         Top             =   2010
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   1402
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
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   780
         Left            =   9435
         Top             =   810
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "ç›å…ãÊï™"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.CheckBox B1Chk_ZKKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "égópíPà "
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
            Left            =   45
            TabIndex        =   17
            Top             =   480
            Width           =   1140
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D21 
         Height          =   750
         Left            =   5280
         Top             =   60
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   1323
         ForeColor       =   16711680
         Caption         =   "çﬁéÌ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_Zaisy 
            Height          =   360
            Left            =   90
            TabIndex        =   7
            Top             =   330
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
         Begin Cis3D_v60.CIS3D B1lb_Zaisy 
            Height          =   345
            Left            =   450
            Top             =   330
            Width           =   1605
            _ExtentX        =   2831
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
      Begin Cis3D_v60.CIS3D CIS3D22 
         Height          =   750
         Left            =   7485
         Top             =   60
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   1323
         ForeColor       =   16711680
         Caption         =   "å`èÛ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_Keijo 
            Height          =   360
            Left            =   90
            TabIndex        =   8
            Top             =   330
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
         Begin Cis3D_v60.CIS3D B1lb_Keijo 
            Height          =   345
            Left            =   450
            Top             =   330
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   780
         Left            =   60
         Top             =   810
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "çﬁéÌÉRÅ[Éh"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_Code 
            Height          =   360
            Left            =   30
            TabIndex        =   10
            Top             =   360
            Width           =   1950
            _ExtentX        =   3440
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
            Text            =   "XXXXXXXXXXXXXXX"
            MaxLength       =   15
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   780
         Index           =   5
         Left            =   6315
         Top             =   810
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "ç›å…î≠íçä«óù"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.ComboBox B1Comb_ZKKbn 
            BackColor       =   &H00C0E0FF&
            Height          =   360
            Left            =   330
            Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
            TabIndex        =   14
            Top             =   345
            Width           =   1365
         End
         Begin CisText_V60.CisText iB1_ZaikoDay 
            Height          =   330
            Left            =   1800
            TabIndex        =   15
            Top             =   360
            Visible         =   0   'False
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   582
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
            Height          =   330
            Left            =   1770
            TabIndex        =   22
            Top             =   360
            Visible         =   0   'False
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   582
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
            Height          =   330
            Left            =   1770
            TabIndex        =   16
            Top             =   360
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   582
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
            Left            =   2490
            TabIndex        =   102
            Top             =   435
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D28 
         Height          =   750
         Left            =   9435
         Top             =   60
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1323
         ForeColor       =   16711680
         Caption         =   "é¿êîãÊï™"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Caption         =   "ÇµÇ»Ç¢"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   9
            Top             =   390
            Width           =   1065
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   780
         Index           =   3
         Left            =   8680
         Top             =   810
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "èäóvó ãÊï™"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.OptionButton B1Op_Syoyo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ç|î¬"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   880
            TabIndex        =   19
            Top             =   330
            Width           =   810
         End
         Begin VB.OptionButton B1Op_Syoyo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ñ≥Å@"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   18
            Top             =   330
            Width           =   810
         End
         Begin VB.OptionButton B1Op_Syoyo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "èç|"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   20
            Top             =   555
            Width           =   810
         End
         Begin VB.OptionButton B1Op_Syoyo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "íYëfç|"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   880
            TabIndex        =   21
            Top             =   555
            Width           =   1050
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
      Height          =   705
      Left            =   240
      ScaleHeight     =   645
      ScaleWidth      =   5760
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   435
      Width           =   5820
      Begin Cis3D_v60.CIS3D BLB_HINBN 
         Height          =   540
         Left            =   90
         Top             =   60
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   953
         BackColor       =   12632319
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
         cAlingnment     =   1
         cPositionX      =   180
         Begin CisText_V60.CisText iH1_Zaist 
            Height          =   360
            Left            =   840
            TabIndex        =   0
            Top             =   120
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
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   540
         Left            =   4200
         Top             =   60
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   953
         BackColor       =   12632319
         ForeColor       =   16711680
         Caption         =   "ê°ñ@"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         cPositionX      =   120
         Begin VB.CheckBox H1Chk_Size 
            BackColor       =   &H00C0FFC0&
            Caption         =   "óL"
            Height          =   285
            Left            =   750
            TabIndex        =   1
            Top             =   150
            Width           =   525
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
      Height          =   705
      Left            =   6390
      ScaleHeight     =   645
      ScaleWidth      =   8595
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   435
      Width           =   8655
      Begin CisBtn_60.CisBtn PB_JIS 
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
         BtnCaption      =   "JIS"
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
         PFForeColor     =   12648447
         PFShadowColor   =   0
         PFBackColor     =   16711935
      End
      Begin CisBtn_60.CisBtn PB_Chg 
         Height          =   585
         Left            =   3465
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
         BtnCaption      =   "éËîzêÊ"
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
      Begin CisBtn_60.CisBtn PB_Hin 
         Height          =   585
         Left            =   4320
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
      Begin CisBtn_60.CisBtn PB_CAN 
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
         Left            =   7725
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
         Left            =   15
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
         Left            =   5160
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         Enabled         =   -1  'True
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
         Left            =   4305
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         Enabled         =   -1  'True
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
      Left            =   960
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   345
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Left            =   930
      Top             =   465
   End
   Begin VB.Timer Timer2 
      Left            =   1200
      Top             =   465
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
      Caption         =   "Åy çﬁóøÉ}ÉXÉ^ÉÅÉìÉeÉiÉìÉX  Åz"
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
         TabIndex        =   97
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
         Left            =   8040
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
   Begin VB.Timer Timer3 
      Left            =   1770
      Top             =   -120
   End
   Begin VB.PictureBox B1_Area2 
      BackColor       =   &H00C0C0C0&
      Height          =   2400
      Left            =   240
      ScaleHeight     =   2340
      ScaleWidth      =   14745
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   3360
      Width           =   14805
      Begin vsFlexLib.vsFlexArray VsGridT 
         Height          =   2220
         Left            =   60
         TabIndex        =   87
         Top             =   60
         Width           =   14595
         _Version        =   196608
         _ExtentX        =   25744
         _ExtentY        =   3916
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   $"CXM0070.frx":0000
         Rows            =   6
         Cols            =   48
         BackColor       =   12648447
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   330
         SelectionMode   =   1
         BackColorAlternate=   12648447
      End
   End
   Begin VB.PictureBox B4_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   5295
      Left            =   11340
      ScaleHeight     =   5235
      ScaleWidth      =   3645
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   5820
      Width           =   3705
      Begin vsFlexLib.vsFlexArray VsGridH 
         Height          =   4620
         Left            =   60
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   540
         Width           =   3510
         _Version        =   196608
         _ExtentX        =   6191
         _ExtentY        =   8149
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   " No|XXXXXXXXX1XXXXXXXXX1X|é–ì‡îwî‘|ÉsÉbÉ`|éÊêî|ñ¢égóp|égópó |óLå¯ä˙å¿|"
         Rows            =   14
         Cols            =   20
         BackColor       =   12648447
         BackColorSel    =   -2147483646
         ForeColorSel    =   16777215
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   330
         SelectionMode   =   1
         AllowUserResizing=   1
         BackColorAlternate=   12648447
      End
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   405
         Left            =   60
         Top             =   90
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   714
         BackColor       =   4210752
         ForeColor       =   16744703
         Caption         =   "égópïiî‘èÓïÒ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
   End
   Begin VB.PictureBox H2_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   2160
      Left            =   240
      ScaleHeight     =   2100
      ScaleWidth      =   3960
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   1170
      Width           =   4020
      Begin VB.PictureBox Picture2 
         Height          =   2070
         Left            =   0
         ScaleHeight     =   2010
         ScaleWidth      =   3870
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   15
         Width           =   3930
         Begin vsFlexLib.vsFlexArray VsGrid1 
            Height          =   1425
            Left            =   30
            TabIndex        =   5
            Top             =   600
            Width           =   3810
            _Version        =   196608
            _ExtentX        =   6720
            _ExtentY        =   2514
            _StockProps     =   228
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "No|###0.00 |###0.00 |###0.00 |çﬁóøä«óùî‘çÜ|çﬁóøñº|çﬁéÌ|å`èÛ|ÉRÅ[Éh|íPà ãÊï™|ä∑éZíl|éËîzíPâø|ç›å…ä«óù|"
            Rows            =   4
            Cols            =   15
            BackColor       =   12648447
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   330
            SelectionMode   =   1
            BackColorAlternate=   12648447
         End
         Begin VB.PictureBox H2LB_Size_B 
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3495
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   -30
            Width           =   3555
            Begin Cis3D_v60.CIS3D CIS3D19 
               Height          =   585
               Left            =   2040
               Top             =   0
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   1032
               BackColor       =   12632319
               ForeColor       =   16711680
               Caption         =   "í∑Ç≥"
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
               Begin CisText_V60.CisText iH2_Long 
                  Height          =   360
                  Left            =   45
                  TabIndex        =   4
                  Top             =   225
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
               Height          =   585
               Left            =   -60
               Top             =   0
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   1032
               BackColor       =   12632319
               ForeColor       =   16711680
               Caption         =   "î¬å˙"
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
               Begin CisText_V60.CisText iH2_Itatu 
                  Height          =   360
                  Left            =   45
                  TabIndex        =   2
                  Top             =   225
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
               Height          =   585
               Left            =   990
               Top             =   0
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   1032
               BackColor       =   12632319
               ForeColor       =   16711680
               Caption         =   "ïù"
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
               Begin CisText_V60.CisText iH2_Width 
                  Height          =   360
                  Left            =   45
                  TabIndex        =   3
                  Top             =   225
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
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   585
            Left            =   30
            Top             =   0
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   1032
            BackColor       =   4210752
            ForeColor       =   12648384
            Caption         =   "ê°ñ@"
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
            cAlingnment     =   6
            cPositionY      =   60
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Cis3D_v60.CIS3D B5_JISNm 
      Height          =   345
      Left            =   4410
      Top             =   1260
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      BackColor       =   16777152
      Caption         =   "JISãKäiÅF"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   16777152
      cBoderColor2    =   16777152
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cFont3D         =   2
      cAlingnment     =   1
      cBoderStyle     =   2
      cPositionX      =   90
   End
   Begin Cis3D_v60.CIS3D B5_JISNo 
      Height          =   345
      Left            =   5520
      Top             =   1260
      Visible         =   0   'False
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   609
      BackColor       =   16777152
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   16777152
      cBoderColor2    =   16777152
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cFont3D         =   2
      cAlingnment     =   1
      cBoderStyle     =   2
      cPositionX      =   90
   End
End
Attribute VB_Name = "CXM0070"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***************************************************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :   çﬁóøÉ}ÉXÉ^ÉÅÉìÉeÉiÉìÉX
'**       ÉtÉHÅ[ÉÄID    :   CXM0070
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :   2003/11/25  By CIS SIMPR-A
'**       ïœ  çX  ì˙    :   2004/04/12  ÉoÉOèCê≥
'**       ïœ  çX  ì˙    :   2004/04/17  ÉoÉOèCê≥
'**       ïœ  çX  ì˙    :   2004/06/21  èCê≥ÅïçÌèúéûÅAê°ñ@ÇÃì¸óÕÇâ¬î\Ç…Ç∑ÇÈ
'**       ïœ  çX  ì˙    :   2004/09/13  î[ïièë/éÛóÃèëå`ë‘ÉRÉìÉ{ãÛîíçsí«â¡
'**       ïœ  çX  ì˙    :   2004/10/27  ÉçÉbÉgÉ`ÉFÉbÉNèCê≥
'**       ïœ  çX  ì˙    :   2007/12/11  By CIS î[ïièëî≠çsãÊï™•éxããèëî≠çsãÊï™•éÛóÃèëî≠çsãÊï™Å@ÇÃí«â¡
'**                                             ÅEî[ïièëå`ë‘•éÛóÃèëå`ë‘Çï\é¶çÌèú(B3Comb_Nouhin,B3Comb_Jyuryo)
'**                                             ÅEñºèÃÉ}ÉXÉ^"ì`ï[éÌóﬁéÛíç•î≠íç"ÇÃâpêîéö1Ç…î[ïièëå`ë‘•éÛóÃèëå`ë‘Çê›íË
'**       ïœ  çX  ì˙    :   2008/02/06  BY CIS SIMPR-AÇÊÇËà⁄çs   êÁë„ìcçHã∆ópÉJÉXÉ^É}ÉCÉY
'**       ïœ  çX  ì˙    :   2008/06/11  BY CIS éËîzãÊï™ÅFéxãã í«â¡
'**       ïœ  çX  ì˙    :   2008/06/27  BY CIS éÛããïiî‘/î≠íçïiî‘ ⁄ÕﬁŸÇéÊà¯êÊÇ…ïœçX/íPâøåvéZéûÇÃä∑éZíl(ì¸óÕñ≥ÇÃèÍçáÇÕÇP)Çí«â¡
'**       ïœ  çX  ì˙    :   2009/01/15  BY CIS íºëóéûÅFíºëóì`ï[î≠çsãÊï™ í«â¡
'**       ïœ  çX  ì˙    :   2009/02/26  BY CIS èäóvó ãÊï™ÅEâ¡çHêÊÅEìæà”êÊÅEé©é–ïiî‘ÅEÇ©ÇÒÇŒÇÒÅEåªïiï[Å@í«â¡
'**       ïœ  çX  ì˙    :   2009/03/02  By CIS èàóùãÊï™ï Ç…òAågçﬁóøÉ}ÉXÉ^Ç÷
'**       ïœ  çX  ì˙    :   2009/04/01  By CIS çﬁóøêVãKÇ…ÇƒçﬁóøéÊà¯êÊòAë±ì¸óÕéûÇÃïsãÔçáèCê≥
'**       ïœ  çX  ì˙    :   2009/05/15  BY CIS (èCê≥)çﬁóøìoò^éûÅAä˘ë∂ê°ñ@ÇëIëíÜÇ…êVãKí«â¡Ç…Ç∑ÇÈÇ∆çﬁóøä«óùî‘çÜÇ™èdï°Ç…Ç»ÇÈ
'**       ïœ  çX  ì˙    :   2009/06/23  BY CIS SKãÊï™ = Ç©ÇÒÇŒÇÒÇÃèÍçáÇ‡îwî‘çÜñ≥ÇnÇjÇ…ïœçX
'**
'**
'**       çÏ  ê¨  ì˙    :   2009/09/23  By CIS SIMPR-A for TOKI
'**       ïœ  çX  ì˙    :   2009/12/08  By CIS ìyäÚí√ópïœçX•••Ç©ÇÒÇŒÇÒêFîpé~ÅAíuèÍédólïœçX
'**       ïœ  çX  ì˙    :   2010/01/16  By CIS ìyäÚí√ópïœçX•••çﬁóøéÊà¯êÊœΩ¿§íIâµä«óùãÊï™í«â¡
'**       ïœ  çX  ì˙    :   2010/03/17  By CIS ìyäÚí√ópïœçX•••ïKê{çÄñ⁄îwåiêFïœçX§éËîzãÊï™ñ¢ì¸óÕéûèâä˙ílæØƒ(∑€)§åüçıâÊñ  ê°ñ@/é©é–ïiî‘í«â¡(â°éRVer)§çﬁóøñºïKê{èåèÇ©ÇÁèúäO
'**
'***************************************************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' Ã´∞—±∏√®Ãﬁ Ã◊∏ﬁ
    Dim ProcHB          As String * 2   ' ì¸óÕàÊêÿë÷
    Dim SyoriKB         As String       ' èàóùãÊï™(êVãK•èCê≥•çÌèú ìô)
    Dim SyoriNM         As String       ' èàóùñºèÃ(êVãK•èCê≥•çÌèú ìô)
    Dim SyoriKB2        As String       ' èàóùãÊï™(êVãK•èCê≥•çÌèú ìô)
    Dim SyoriNM2        As String       ' èàóùñºèÃ(êVãK•èCê≥•çÌèú ìô)
    Dim SyoriKB3        As String       ' èàóùãÊï™(êVãK•èCê≥•çÌèú ìô)
    Dim SyoriNM3        As String       ' èàóùñºèÃ(êVãK•èCê≥•çÌèú ìô)
    
    Dim mCHK            As Boolean
    Dim RowColFlg       As Boolean
    Dim GridBack        As String
    
    Dim DayString()     As String
    
    Dim mEdit           As String

    Dim mGotColor       As Long
    Dim mLostColor      As Long
    
    Dim CisVsGrid       As New CisVsGrid3   ' ê°ñ@èÓïÒ
    Dim CisVsGridT      As New CisVsGrid3   ' éËîzèÓïÒ
    Dim CisVsGridH      As New CisVsGrid3   ' égópïiî‘
    Dim CisComboTH      As CisCombo
    Dim CisComboZK      As CisCombo
'''2007/12/11DEL    Dim CisComboNH      As CisCombo
'''2007/12/11DEL    Dim CisComboJR      As CisCombo
    Dim CisComboMS      As CisCombo
        
    Dim mIx             As Long
    Dim mIx2            As Long
    Dim mStr            As String
    
    Dim ActObj          As Object
    Dim ActRow          As Integer
    
    Dim KCol            As Long         ' îÒï\é¶çÄñ⁄äJén∂◊—
    
    Dim SCol()          As Long         ' ï\é¶çÄñ⁄∂◊—
    Dim SColT()         As Long         ' ï\é¶çÄñ⁄∂◊—
    Dim SColH()         As Long         ' ï\é¶çÄñ⁄∂◊—
    
    Dim wKanri          As Long         ' çﬁóøä«óùî‘çÜ
    Dim wsDbl           As Double
    Dim wsVsGridT_Row   As Long
    Dim wsVsGridH_Row   As Long

    Dim wsROW           As Long
'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2007/12/11 Update Satart
    Dim SaveKeitai      As String
'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2007/12/11 Update End

    Dim mBackColor      As Long         ' îwåiêFÉfÉtÉHÉãÉg  ' 2009.02.27 add

'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås    ÇkÇèÇÅÇÑÉCÉxÉìÉg    Åt              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    If Rv_Call3 = "" Then
    '   #------------------#
    '   # ìÒ èd ãN ìÆ ñh é~ #
    '   #------------------#
        If Not Dupli_Start Then End
    '   #------------------#
    '   # èâ ä˙ ì‡ óe éÊ ìæ #
    '   #------------------#
        If Not IniGet Then End
    End If
'   #------------------#
'   #  çÄ ñ⁄ èâ ä˙ âª   #
'   #------------------#
    Set CisFrm = New CISFormContorl
    
'-------------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA start
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE START
''    mGotColor = H1Chk_Size.BackColor
''    mLostColor = B3Op_SKKbn(1).BackColor
'    mGotColor = B2Op_Hinmoku(1).BackColor
'    mLostColor = B2Op_Hinmoku(0).BackColor
'    B2Op_Hinmoku(1).BackColor = mLostColor
'    B2Op_Hinmoku(1).Value = True
'    Call B2Op_Hinmoku_Click(1)
'    B2Op_Hinmoku(1).BackColor = mLostColor
'    B2Op_Hinmoku(1).Value = False
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE END
'-------------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA end
'-------------------------------------------------------------------------------' 2009.02.27 add : ïiñ⁄ä÷òA start
    mGotColor = H1Chk_Size.BackColor
    mLostColor = B3Op_SKKbn(1).BackColor
    B2Op_Hinmoku(0).Tag = 1
    B2Op_Hinmoku(1).Visible = True
'-------------------------------------------------------------------------------' 2009.02.27 add : ïiñ⁄ä÷òA end
    mBackColor = &HFFFFFF       ' îwåiêFÉfÉtÉHÉãÉg                              ' 2009.02.27 add

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
'    Dummy2.Left = -1000
'    Dummy3.Left = -1000
    Timer1.Interval = 1000   ' éûçèï\é¶ä‘äu
    Timer1.Enabled = True    ' éûçèï\é¶äJén
    Timer2.Interval = 500    ' èàóùãÊï™ì_ñ≈ä‘äu(äÓñ{)
    Timer2.Enabled = False   ' èàóùãÊï™ì_ñ≈ã÷é~(äÓñ{)
'    Timer3.Interval = 500    ' èàóùãÊï™ì_ñ≈ä‘äu(éÊà¯êÊ)
'    Timer3.Enabled = False   ' èàóùãÊï™ì_ñ≈ã÷é~(éÊà¯êÊ)
    
    ' ñºèÃÉ}ÉXÉ^ämîF
    Call MeisyoInit
    '+--------------------------------+
    '+  ÉOÉäÉbÉhèâä˙ê›íË(çﬁóøèÓïÒ)
    '+--------------------------------+
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 350
        .PatanMax = 1
        .SelectCol = 6
        .InitGet = False
        .Init
    End With
    Call GridInit
    '+--------------------------------+
    '+  ÉOÉäÉbÉhèâä˙ê›íË(éËîzèÓïÒ)
    '+--------------------------------+
    With CisVsGridT
        Set .GridObj = VsGridT
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 350
        .PatanMax = 1
        .SelectCol = 45
        .InitGet = False
        .Init
    End With
    Call GridInitT
    '+--------------------------------+
    '+  ÉOÉäÉbÉhèâä˙ê›íË(égópïiî‘èÓïÒ)
    '+--------------------------------+
    With CisVsGridH
        Set .GridObj = VsGridH
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 350
        .PatanMax = 1
        .SelectCol = 18
        .InitGet = False
        .Init
    End With
    Call GridInitH
    '+---------------------+
    '+ ∫›ŒﬁŒﬁØ∏Ω èâä˙ê›íË  +
    '+---------------------+
    ' éËîzãÊï™
    Set CisComboTH = New CisCombo
    With CisComboTH
        Set .Connect = CisDB
        Set .CombObj = B3Comb_THKbn
        .Control = ""
        .ComboName = "éËîzãÊï™"
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .Ess = True
        .ComboInit
        
        '=======================================***** 2008/06/11 Start
        '.CombowkSet (3)
        .CombowkSet (4)
        '=======================================***** 2008/06/11 End
        .Code = 0: .CombText = "çwì¸"
        .CombowkSet
        .Code = 1: .CombText = "íºëó"
        .CombowkSet
        .Code = 3: .CombText = "éÛãã"
        .CombowkSet
        '======================================***** 2008/06/11 Start
        .Code = 4: .CombText = "éxãã"
        .CombowkSet
        '======================================***** 2008/06/11 End
    End With
    ' ç›å…ãÊï™
    Set CisComboZK = New CisCombo
    With CisComboZK
        Set .Connect = CisDB
        Set .CombObj = B1Comb_ZKKbn
        .Control = "ñºèÃÉ}ÉXÉ^"
        .ComboName = "ç›å…ãÊï™"
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 15
        .Ess = True
        .ComboInit
    End With
    
    ' î[ïièëå`ë‘
'''2007/12/11DEL    Set CisComboNH = New CisCombo
'''2007/12/11DEL    With CisComboNH
'''2007/12/11DEL        Set .Connect = CisDB
'''2007/12/11DEL        Set .CombObj = B3Comb_Nohin
'''2007/12/11DEL        .Control = "ñºèÃÉ}ÉXÉ^"
'''2007/12/11DEL        .ComboName = "î[ïièëå`ë‘î≠íç"
'''2007/12/11DEL        .IniPath = gIniExe
'''2007/12/11DEL        .SetSpace = 50
'''2007/12/11DEL        .Ess = False
'''2007/12/11DEL        .ComboInit
'''2007/12/11DEL    End With
'''2007/12/11DEL    ' éÛóÃèëå`ë‘
'''2007/12/11DEL    Set CisComboJR = New CisCombo
'''2007/12/11DEL    With CisComboJR
'''2007/12/11DEL        Set .Connect = CisDB
'''2007/12/11DEL        Set .CombObj = B3Comb_Jyuro
'''2007/12/11DEL        .Control = "ñºèÃÉ}ÉXÉ^"
'''2007/12/11DEL        .ComboName = "éÛóÃèëå`ë‘"
'''2007/12/11DEL        .IniPath = gIniExe
'''2007/12/11DEL        .SetSpace = 50
'''2007/12/11DEL        .Ess = False
'''2007/12/11DEL        .ComboInit
'''2007/12/11DEL    End With
    ' ñæç◊èëî≠çs
    Set CisComboMS = New CisCombo
    With CisComboMS
        Set .Connect = CisDB
        Set .CombObj = B3Comb_Meisi
        .Control = "ñºèÃÉ}ÉXÉ^"
        .ComboName = "ñæç◊èëî≠çs"
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .Ess = True
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

    If Rv_Call3 = "" Then
        Me.Show
    End If
End Sub
'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    ÅsÇ`ÇÉÇîÇâÇñÇÅÇîÇÖ ÉCÉxÉìÉgÅt                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
' *ADD* 2004/04/17
    Call CisComboZK.ComboSet
    With CisComboZK
        .Code = ZRM.ç›å…ä«óùãÊï™
        .TextGet
    End With

    If Rv_Call3 <> "" Then
        Call Disp_Call
        PB_END.BtnCaption = "ñﬂÇÈ"
    End If

    FormAct = Not FormAct
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
                    Call PB_ADD_Click    'ÅyêVãKÅz
                End If
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    'ÅyèCê≥Åz
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    'ÅyçÌèúÅz
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  'ÅyåüçıÅz
           Case vbKeyF9:       If PB_CHG.Visible Then Call PB_CHG_Click    'ÅyéËîzêÊÅz
           Case vbKeyF10:      If PB_Hin.Visible Then Call PB_Hin_Click    'Åyégópïiî‘Åz
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  'Åyå„ëﬁÅz
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  'ÅyëOêiÅz
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    'ÅyéÊè¡Åz
           Case vbKeyF11:      If PB_JIS.Visible Then Call PB_JIS_Click    'ÅyJISÅz
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    'ÅyèIóπÅz
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      'Åyì¸óÕÅz
                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
                                  Dummy.Enabled = False
                               End If
'-------------------------------------------------------------------------------' 2009.02.27 del start
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
'           Case vbKeyF5:       If PB_KTN.Visible Then Call PB_KTN_Click    'Åyçwì¸íPâøÅz
'           Case vbKeyF6:       If PB_STN.Visible Then Call PB_STN_Click    'ÅyéxããíPâøÅz
'           'Case vbKeyF11:      MsgBox (ActiveControl.Name)
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'-------------------------------------------------------------------------------' 2009.02.27 del end
           Case Else:          Exit Function
    End Select
    Key_Acc = True
End Function
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås QueryUnloadÉCÉxÉìÉg Åt                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    If Rv_Call3 = "" Then
       Cancel = 1
    End If
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      åƒèoÇ≥ÇÍÇΩèÍçáÇÃèâä˙ï\é¶
'+----------------------------------------------------------+
Private Sub Disp_Call()
    iH1_Zaist = RV_Zaist
    Call ReturnPress
    
    SyoriKB = "MNT"
    Call SyoriName(SyoriKB)
    
    vsGrid1.Row = wsROW
    Call VSGrid1_ClickClick
    Call CisVsGrid.vsColor(å¬ï )
    If Not vsGrid1.RowIsVisible(wsROW) Then
        vsGrid1.TopRow = wsROW
    End If
    
    PB_CHG.Tag = "TOR"
    Call PB_CHG_Click
    PB_CHG.Visible = False
End Sub

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
        Call CisFrm.UnLoadDisp
        Unload CXM0070
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
   
    If ProcHB = "H2" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("H2")
        Call HeadBodyClear("B")
        Call SyoriName(SyoriKB)
        
        iH1_Zaist.SetFocus
        GoTo PB_CAN_Ed
    End If
        
    If ProcHB = "B1" Then
        Call HeadBodyClear("B")
        If H1Chk_Size.Value = 0 Then
            ProcHB = "H2"
        Else
            ProcHB = "H1"
        End If
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        
        If SyoriKB = "ADD" Then
            If H1Chk_Size.Value = 0 Then
                iH2_Itatu.SetFocus
            Else
                iH1_Zaist.SetFocus
            End If
        Else
            If H1Chk_Size.Value = 0 Then
                vsGrid1.SetFocus
                Call VSGrid1_ClickClick
            Else
                iH1_Zaist.SetFocus
            End If
            
        End If
        GoTo PB_CAN_Ed
    End If
' B2
    If ProcHB = "B2" Then
        If Rv_Call3 <> "" Then
            Call PB_END_Click
            Exit Sub
        End If
        
        SyoriKB2 = ""
        Call SyoriName2(SyoriKB2)
        ProcHB = "B1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("B2")
        Call HeadBodyClear("B3")
        Call HeadBodyClear("B6")
        Call HeadBodyClear("B7")
        Call HeadBodyClear("B8")
        Call HeadBodyClear("B4")            ' 2009.03.03 add
        
        iB1_Name.SetFocus
        GoTo PB_CAN_Ed
    End If
' B3
   If ProcHB = "B3" Then
      RowColFlg = True
      Call CisVsGridT.vsColor(è¡ãéëSïî)
      Call CisVsGridH.vsColor(è¡ãéëSïî)     ' 2009.03.03 add
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
         DoEvents
         VsGridT.SetFocus
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
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Or ProcHB = "H2" Then
        If SyoriKB <> "ADD" Then
            SyoriKB = "ADD"
            Call SyoriName(SyoriKB)
        End If
        If ProcHB = "H1" Then
            iH1_Zaist.SetFocus
        Else
            Call HeadBodyClear("B1")
            iH2_Itatu = ""
            iH2_Width = ""
            iH2_Long = ""
            iH2_Itatu.SetFocus
        End If
        GoTo PB_ADD_ED
    End If
    
    If ProcHB = "B2" Or ProcHB = "B3" Then
        If SyoriKB2 <> "ADD" Then
            SyoriKB2 = "ADD"
            Call SyoriName2(SyoriKB2)
        End If
        iB2_Torcd.SetFocus
    
        GoTo PB_ADD_ED
    End If
    
PB_ADD_ED:
    Call CisFrm.MousePT(1)
End Sub

'+----------------------------+
'+    èC  ê≥  ÉL Å[(F2)        +
'+----------------------------+
Private Sub PB_MNT_Click()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Or ProcHB = "H2" Then
        If SyoriKB <> "MNT" Then
            SyoriKB = "MNT"
            Call SyoriName(SyoriKB)
        End If
        If ProcHB = "H1" Then
            iH1_Zaist.SetFocus
        Else
            vsGrid1.SetFocus
            Call VSGrid1_ClickClick
        End If
        GoTo PB_MNT_ED
    End If
    
    If ProcHB = "B2" Or ProcHB = "B3" Then
        If SyoriKB2 <> "MNT" Then
            SyoriKB2 = "MNT"
            Call SyoriName2(SyoriKB2)
        End If
        iB2_Torcd.SetFocus
        GoTo PB_MNT_ED
    End If

PB_MNT_ED:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    çÌ  èú  ÉL Å[(F3)        +
'+----------------------------+
Private Sub PB_DEL_Click()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Or ProcHB = "H2" Then
        If SyoriKB <> "DEL" Then
            SyoriKB = "DEL"
        Else
            SyoriKB = ""
        End If
        Call SyoriName(SyoriKB)
        If ProcHB = "H1" Then
            iH1_Zaist.SetFocus
        Else
            vsGrid1.SetFocus
            Call VSGrid1_ClickClick
        End If
        GoTo PB_DEL_ED
    End If
    
    If ProcHB = "B2" Or ProcHB = "B3" Then
        If SyoriKB2 <> "DEL" Then
            SyoriKB2 = "DEL"
            Call SyoriName2(SyoriKB2)
        End If
        iB2_Torcd.SetFocus
        GoTo PB_DEL_ED
    End If

PB_DEL_ED:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    å„  ëﬁ  ÉL Å[(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = " SELECT çﬁéø,ê°ñ@ãÊï™ "
    gSL_Select = gSL_Select & " FROM çﬁóøÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " WHERE çﬁéø < '" & RTrim(iH1_Zaist) & "'"
    gSL_Select = gSL_Select & " GROUP BY çﬁéø,ê°ñ@ãÊï™"
    gSL_Select = gSL_Select & " ORDER BY çﬁéø Desc "
    Call Head2Input("Back")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Zaist.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    ëO  êi  ÉL Å[(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = " SELECT çﬁéø,ê°ñ@ãÊï™ "
    gSL_Select = gSL_Select & " FROM çﬁóøÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " WHERE çﬁéø > '" & RTrim(iH1_Zaist) & "'"
    gSL_Select = gSL_Select & " GROUP BY çﬁéø,ê°ñ@ãÊï™"
    gSL_Select = gSL_Select & " ORDER BY çﬁéø "
    Call Head2Input("Next")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Zaist.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+--------------------------------------------+
'+    äÓñ{ÅEéËîzêÊ(ïiî‘)êÿë÷ÉLÅ[(F9)
'+--------------------------------------------+
Private Sub PB_CHG_Click()
    If PB_CHG.Tag = "TOR" Then
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE START
'        Call HeadBodyClear("B2")
'        Call HeadBodyClear("B3")
'        ProcHB = "B2"
'        Call DispChange(ProcHB)
'        SyoriKB2 = "ADD"
'        Call SyoriName2(SyoriKB2)
'        iB2_Torcd.SetFocus
        If Not Body1Chk Then Exit Sub
        Call HeadBodyClear("B2")
        Call HeadBodyClear("B3")
'-----------------------------------------------' 2009.03.03 add start
        RowColFlg = True
        Call CisVsGridH.vsColor(è¡ãéëSïî)
        RowColFlg = False
'-----------------------------------------------' 2009.03.03 add end
        ProcHB = "B2"
        Call DispChange(ProcHB)
        SyoriKB2 = "ADD"
        Call SyoriName2(SyoriKB2)
        B2Op_Hinmoku(1).Value = True
        iB2_Torcd.SetFocus
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE END
    Else
        Call PB_CAN_Click
    End If
End Sub
'+-------------------------------------------+
'+    çﬁóøÅEïiî‘ä÷òAïtÇØ (PF10)
'+-------------------------------------------+
Private Sub PB_Hin_Click()

    If Not Body1Chk Then Exit Sub
    If Not DBPut(False) Then Exit Sub
    
    RV_Left = 0
    RV_Top = 0
    Rv_Call2 = "ZAI"
    RV_Zaist = RTrim(iH1_Zaist)
    RV_Itatu = CisFun.Val2(iH2_Itatu)
    RV_Width = CisFun.Val2(iH2_Width)
    RV_Long = CisFun.Val2(iH2_Long)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
    RV_Tani = RTrim(B1lb_Tani)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END

    CXM0110.Show vbModal
    Unload CXM0110
    Set CXM0110 = Nothing

    Call GridSetH
End Sub

'-------------------------------------------------------' 2009.09.23 add start
'+-------------------------------------------+
'+    JISãKäi (PF11)
'+-------------------------------------------+
Private Sub PB_JIS_Click()
    RV_JIS_Zai = Trim(iH1_Zaist)
    RV_JIS_Num = Trim(B5_JISNo)
        
    CXM0310.Show vbModal
    Unload CXM0310
    Set CXM0310 = Nothing
    
    B5_JISNo.Caption = RV_JIS_Num

End Sub
'-------------------------------------------------------' 2009.09.23 add end

'+----------------------------+
'+    åü  çı  ÉL Å[(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Zaist" Then Call Look_Zaist: Exit Sub
    If PB_Look.Tag = "iB1_Zaisy" Then Call Look_Kubun(0): Exit Sub
    If PB_Look.Tag = "iB1_Keijo" Then Call Look_Kubun(1): Exit Sub
    If PB_Look.Tag = "iB1_Tani" Then Call Look_Kubun(2): Exit Sub
    If PB_Look.Tag = "iB1_TTani" Then Call Look_Kubun(3): Exit Sub
    If PB_Look.Tag = "iB2_Torcd" Then Call Look_Torcd(0): Exit Sub
    If PB_Look.Tag = "iB2_Ukeir" Then Call Look_Ukeir(0): Exit Sub
    If PB_Look.Tag = "iB3_Hako1" Then Call Look_Kubun(4): Exit Sub
    If PB_Look.Tag = "iB3_Hako2" Then Call Look_Kubun(5): Exit Sub
    If PB_Look.Tag = "iB3_KTorcd" Then Call Look_Torcd(1): Exit Sub
    If PB_Look.Tag = "iB3_MTorcd" Then Call Look_Torcd(2): Exit Sub
    If PB_Look.Tag = "iB3_MUkeir" Then Call Look_Ukeir(2): Exit Sub
'-----------------------------------------------------------------------' 2009.02.26 add start
    If PB_Look.Tag = "iB3_STorcd" Then Call Look_Torcd(3): Exit Sub
    If PB_Look.Tag = "iB3_TTorcd" Then Call Look_Torcd(4): Exit Sub
'-----------------------------------------------------------------------' 2009.02.26 add end
    If PB_Look.Tag = "iB3_Densy" Then Call Look_Kubun(6): Exit Sub
    If PB_Look.Tag = "iB3_DensyJ" Then Call Look_Kubun(7): Exit Sub
    
    If PB_Look.Tag = "iB3_CDate" Then
        iB3_CDate.ShowCalender
        Exit Sub
    End If
    If PB_Look.Tag = "iB3_SDate" Then
        iB3_SDate.ShowCalender
        Exit Sub
    End If
    If PB_Look.Tag = "iB3_EDate" Then
        iB3_EDate.ShowCalender
        Exit Sub
    End If

    If PB_Look.Tag = "iB3_Okiba1" Then Call Look_Kubun(13): Exit Sub           '2009/12/08
    If PB_Look.Tag = "iB3_Okiba2" Then Call Look_Kubun(14): Exit Sub           '2009/12/08
End Sub
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
'+----------------------------+
'+       ïiî‘åüçı             +
'+      KB = 0 : éÛíçïiî‘     +
'+           1 : î≠íçïiî‘     +
'+----------------------------+
Private Sub Look_Hinbn(KB As Integer)

    RV_Left = 0
    RV_Top = 0
    If KB = 0 Then
       RVI_Hinmoku = 0
    Else
       RVI_Hinmoku = 1
    End If
    RVI_SK = 9
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    '=============================================== *** 2008/06/23 Start
'    If RV_Rtn Then
'       If KB = 0 Then
'          iB1_JHinbn = RV_Hinbn
'       Else
'          iB1_HHinbn = RV_Hinbn
'       End If
'    End If
    
    If RV_Rtn Then
       If KB = 0 Then
          iB3_JHinbn = RV_Hinbn
       Else
          iB3_HHinbn = RV_Hinbn
       End If
    End If
    '=============================================== *** 2008/06/23 End

End Sub
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'+----------------------------+
'+       çﬁóøåüçı              +
'+----------------------------+
Private Sub Look_Zaist()
    
'    RV_Left = 0
'    RV_Top = 0
'    RV_Call = "C"

    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 9
    RVI_SK = 9

'    CKK0075.Show vbModal
'    Unload CKK0075
'    Set CKK0075 = Nothing
    CKK0077.Show vbModal
    Unload CKK0077
    Set CKK0077 = Nothing

    If RV_Rtn Then
       
        iH1_Zaist = RV_Zaist
    
        Call Head2Input("INV")
        Dummy.Enabled = True: Dummy.SetFocus
        iH1_Zaist.SetFocus
        Dummy.Enabled = False
    End If

End Sub
'+----------------------------+
'+       éÊà¯êÊåüçı            +
'+----------------------------+
Private Sub Look_Torcd(KB As Integer)

    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    B2Op_Hinmoku(1).Value = True                                                    ' 2009.02.27 add : ïiñ⁄ä÷òA
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE START
'    Select Case KB
'        Case 0: RV_TorKb = 9
'        Case 1: RV_TorKb = 5
'        Case 2: RV_TorKb = 8
'    End Select
    If KB = 0 And B2Op_Hinmoku(0).Value Then
        RV_TorKb = 3
    End If
    If KB = 0 And B2Op_Hinmoku(1).Value Then
        RV_TorKb = 9
    End If
    If KB = 1 And B2Op_Hinmoku(0).Value Then
        RV_TorKb = 2
    End If
    If KB = 1 And B2Op_Hinmoku(1).Value Then
        RV_TorKb = 5
    End If
    If KB = 2 And B2Op_Hinmoku(0).Value Then
        RV_TorKb = 3
    End If
    If KB = 2 And B2Op_Hinmoku(1).Value Then
        RV_TorKb = 8
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE END
'---------------------------------------------------------------' 2009.02.26 add start
    If KB = 3 And B2Op_Hinmoku(0).Value Then    ' â¡çHêÊ
        RV_TorKb = 7 '3
    End If
    If KB = 3 And B2Op_Hinmoku(1).Value Then
        RV_TorKb = 7
    End If
    If KB = 4 And B2Op_Hinmoku(0).Value Then    ' ìæà”êÊ
        RV_TorKb = 2 '3
    End If
    If KB = 4 And B2Op_Hinmoku(1).Value Then
        RV_TorKb = 2
    End If
'---------------------------------------------------------------' 2009.02.26 add end
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
       
        Select Case KB
            Case 0
                iB2_Torcd = RV_TorcdK
                B2lb_Torcd = RV_TorRName
            Case 1
                iB3_KTorcd = RV_TorcdK
                B3lb_KTorcd = RV_TorRName
            Case 2
                iB3_MTorcd = RV_TorcdK
                B3lb_MTorcd = RV_TorRName
'---------------------------------------------------------------' 2009.02.26 add start
            Case 3  ' â¡çHêÊ
                iB3_STorcd = RV_TorcdK
                B3lb_STorcd = RV_TorRName
            Case 4  ' ìæà”êÊ
                iB3_TTorcd = RV_TorcdK
                B3lb_TTorcd = RV_TorRName
'---------------------------------------------------------------' 2009.02.26 add end
        End Select
    End If
End Sub
'+-----------------------------------------------------------------------------+
'+       éÊà¯êÊéÛì¸åüçı
'+-----------------------------------------------------------------------------+
Private Sub Look_Ukeir(KB As Integer)
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    If KB = 0 Then
        B2Op_Hinmoku(1).Value = True                                                ' 2009.02.27 add : ïiñ⁄ä÷òA
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE START
        'RV_TorKb = 9
        If B2Op_Hinmoku(0).Value Then
           RV_TorKb = 3
        Else
           RV_TorKb = 9
        End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE END
        RV_Torcd = iB2_Torcd
    Else
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE START
        'RV_TorKb = 8
        If B2Op_Hinmoku(0).Value Then
           RV_TorKb = 3
        Else
           RV_TorKb = 9
        End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE END
        RV_Torcd = iB3_MTorcd
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
            iB3_MTorcd = RV_TorcdK
            iB3_MUkeir = RV_Ukeir
            B3lb_MTorcd = RV_TorRName
        End If
    End If
End Sub
'+-----------------------------------------------------------------------------+
'+       ñºèÃÉ}ÉXÉ^åüçı
'+          KB  : 0     çﬁéÌãÊï™(çﬁéÌ)
'+              : 1     å`èÛãÊï™(å`èÛ)
'+              : 2     íPà ãÊï™(íPà )
'+              : 3     íPà ãÊï™(éËîzíPà )
'+              : 4     é˚óeäÌä«óù(é˚óeäÌá@)
'+              : 5     é˚óeäÌä«óù(é˚óeäÌáA)
'+              : 6     ì`ï[éÌóﬁî≠íç(ì`ï[éÌóﬁ)
'+              : 7     ì`ï[éÌóﬁéÛíç(ì`ï[éÌóﬁ)
'+-----------------------------------------------------------------------------+
Private Sub Look_Kubun(KB As Integer)

    RV_Left = 0
    RV_Top = 0

    Select Case KB
        Case 0: RVI_Kubun = "çﬁéÌãÊï™"
        Case 1: RVI_Kubun = "å`èÛãÊï™"
        Case 2: RVI_Kubun = "íPà ãÊï™"
        Case 3: RVI_Kubun = "íPà ãÊï™"
        Case 4: RVI_Kubun = "é˚óeäÌä«óù"
        Case 5: RVI_Kubun = "é˚óeäÌä«óù"
        Case 6: RVI_Kubun = "ì`ï[éÌóﬁî≠íç"
        Case 7: RVI_Kubun = "ì`ï[éÌóﬁéÛíç"      ' 2008/02/06 ADD
        Case 13: RVI_Kubun = "ä«óùèÍèä"             '2009/12/08
        Case 14: RVI_Kubun = "ä«óùèÍèäÉÜÅ[ÉUÅ["     '2009/12/08
        Case Else: Exit Sub
    End Select

    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
        Select Case KB
            Case 0
                iB1_Zaisy = RV_Code
                Call iB1_Zaisy_LostFocus
            Case 1
                iB1_Keijo = RV_Code
                Call iB1_Keijo_LostFocus
            Case 2
                iB1_Tani = RV_Code
                Call iB1_Tani_LostFocus
            Case 3
                iB1_TTani = RV_Code
                Call iB1_TTani_LostFocus
            Case 4
                iB3_Hako1 = RV_Code
            Case 5
                iB3_Hako2 = RV_Code
            Case 6
                iB3_Densy = RV_Code
                 Call iB3_Densy_LostFocus
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
            Case 7
                iB3_DensyJ = RV_Code
                 Call iB3_DensyJ_LostFocus
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
            '=================================== 2009/12/08 Start
            Case 13
                iB3_Okiba1 = RV_Code
                Call iB3_Okiba1_LostFocus
            Case 14
                iB3_Okiba2 = RV_Code
                Call iB3_Okiba2_LostFocus
            '=================================== 2009/12/08 End
        End Select
    End If

End Sub
'---------------------------------------------------------------------------------------' 2009.02.27 del : Åiïiñ⁄ä÷òAä‹ÇﬁÅjstart
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
''+--------------------------------+
''+     çwì¸íPâøÉ{É^ÉìÉNÉäÉbÉN     +
''+--------------------------------+
'Private Sub PB_KTN_Click()
'    Rv_Call3 = "X"
'    RV_HinZai = 1                   ' ïiî‘,çﬁóøãÊï™
'    ' ======================================= *** 2008/06/27 Start
'    'RV_Hinbn = RTrim(iB1_HHinbn)
'    RV_Hinbn = RTrim(iB3_HHinbn)
'    ' ======================================= *** 2008/06/27 End
'
'    Call CisComboTH.CodeGet
'    RV_Torcd = RTrim(iB2_Torcd)
'    ' ======================================= *** 2008/06/27 Start(Comment)
'    'RV_KTorcd = RTrim(iB3_KTorcd)
'    ' ======================================= *** 2008/06/27 Start(Comment)
'    RV_SYmd = ""
'    RV_EYmd = ""
'    Rv_Call3Rtn = 0
'    CXM0180.Show vbModal
'    Unload CXM0180
'    Set CXM0180 = Nothing
'
'    If RV_Rtn Then
'    End If
'
'    Rv_Call3 = ""
'End Sub
''+--------------------------------+
''+     éxããíPâøÉ{É^ÉìÉNÉäÉbÉN     +
''+--------------------------------+
'Private Sub PB_STN_Click()
'    Rv_Call3 = "X"
'    RV_HinZai = 1                   ' ïiî‘,çﬁóøãÊï™
'    '===================================== *** 2008/06/27 Start
''    If B2Op_Hinmoku(0).Value Then
''       RV_Hinbn = RTrim(iB1_JHinbn)
''    Else
''       RV_Hinbn = RTrim(iB1_HHinbn)
''    End If
'
'    If B2Op_Hinmoku(0).Value Then
'       RV_Hinbn = RTrim(iB3_JHinbn)
'    Else
'       RV_Hinbn = RTrim(iB3_HHinbn)
'    End If
'    '===================================== *** 2008/06/27 End
'
'    Call CisComboTH.CodeGet
'    If B2Op_Hinmoku(1).Value And CisComboTH.Code = "1" Then
'       RV_Torcd = RTrim(iB3_MTorcd)
'    Else
'       RV_Torcd = RTrim(iB2_Torcd)
'    End If
'    RV_SYmd = ""
'    RV_EYmd = ""
'    CXM0190.Show vbModal
'    Unload CXM0190
'    Set CXM0190 = Nothing
'
'    If RV_Rtn Then
'    End If
'
'    Rv_Call3 = ""
'End Sub
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'---------------------------------------------------------------------------------------' 2009.02.27 del : Åiïiñ⁄ä÷òAä‹ÇﬁÅjend
'+---( H1 )----------------------------------------------------------------------------
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
'+  H1Chk_Size
'+-----------------------------+
Private Sub H1Chk_Size_Click()
    If H1Chk_Size.Value = 0 Then
        H1Chk_Size.BackColor = mGotColor
        H1Chk_Size.Caption = "óL"
    Else
        H1Chk_Size.BackColor = mLostColor
        H1Chk_Size.Caption = "ñ≥"
    End If
End Sub
'+---( B1 )----------------------------------------------------------------------------
'+-----------------------------+
'+  iB1_Zaisy
'+-----------------------------+
Private Sub iB1_Zaisy_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Zaisy_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    
    B1lb_Zaisy = ""
    If MeisyoGet("çﬁéÌãÊï™", iB1_Zaisy) Then
        B1lb_Zaisy = RTrim(SYM_Meisyo)
    End If
End Sub
'+-----------------------------+
'+  iB1_Keijo
'+-----------------------------+
Private Sub iB1_Keijo_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Keijo_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    B1lb_Keijo = ""
    If MeisyoGet("å`èÛãÊï™", iB1_Keijo) Then
        B1lb_Keijo = RTrim(SYM_Meisyo)
    End If

End Sub
'+-----------------------------+
'+  iB1_Tani
'+-----------------------------+
Private Sub iB1_Tani_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Tani_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B1lb_Tani = ""
    If MeisyoGet("íPà ãÊï™", iB1_Tani) Then
        B1lb_Tani = SYM_Meisyo
    End If
End Sub
'+-----------------------------+
'+  iB1_TTani
'+-----------------------------+
Private Sub iB1_TTani_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_TTani_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B1lb_TTani = ""
    If MeisyoGet("íPà ãÊï™", iB1_TTani) Then
        B1lb_TTani = SYM_Meisyo
    End If
End Sub
'+-----------------------------+
'+  B1Chk_JSKbn
'+-----------------------------+
Private Sub B1Chk_JSKbn_Click()
    If B1Chk_JSKbn.Value = 0 Then
        B1Chk_JSKbn.Caption = "ÇµÇ»Ç¢"
    Else
        B1Chk_JSKbn.Caption = "Ç∑ÇÈ"
    End If
End Sub
'+-----------------------------+
'+  B1Comb_ZKKbn
'+-----------------------------+
Private Sub B1Comb_ZKKbn_GotFocus()
    B1Comb_ZKKbn.BackColor = gIGotSel
End Sub
Private Sub B1Comb_ZKKbn_LostFocus()
    B1Comb_ZKKbn.BackColor = gILostSel
End Sub
Private Sub B1Comb_ZKKbn_Click()
    CisComboZK.CodeGet
    gStr = CisComboZK.Code
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
'+  B1Chk_ZKKbn
'+-----------------------------+
Private Sub B1Chk_ZKKbn_Click()
    If B1Chk_ZKKbn.Value = 0 Then
        B1Chk_ZKKbn.Caption = "égópíPà "
    Else
        B1Chk_ZKKbn.Caption = "éËîzíPà "
    End If
End Sub
'-------------------------------------------------------' 2009.02.26 add end
'+-----------------------------+
'+  B1Op_Syoyo_Click
'+-----------------------------+
Private Sub B1Op_Syoyo_Click(INDEX As Integer)
    For gInt = 0 To 3
        B1Op_Syoyo(gInt).BackColor = mLostColor
    Next
    B1Op_Syoyo(INDEX).BackColor = mGotColor
    B1Op_Syoyo(0).Tag = INDEX
End Sub
'-------------------------------------------------------' 2009.02.26 add end

'+---( B2 )----------------------------------------------------------------------------
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
'---------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA start
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE START
'    If B2Op_Hinmoku(0).Value Then gInt = 3 Else gInt = 9
'    'If TorNmGet(iB2_Torcd, 9) Then
'    If TorNmGet(iB2_Torcd, gInt) Then
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE END
'---------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA end
    If TorNmGet(iB2_Torcd, 9) Then                                          ' 2009.02.27 add : ïiñ⁄ä÷òA
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

'---------------------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA start
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
''+-----------------------------+
''+  B2Op_Hinmoku
''+-----------------------------+
'Private Sub B2Op_Hinmoku_Click(INDEX As Integer)
'    If INDEX = 0 Then
'       B2Op_Hinmoku(0).BackColor = mGotColor
'       B2Op_Hinmoku(1).BackColor = mLostColor
'       Back_Seban.Caption = "ìæîwî‘"
'    Else
'       B2Op_Hinmoku(0).BackColor = mLostColor
'       B2Op_Hinmoku(1).BackColor = mGotColor
'       Back_Seban.Caption = "îwî‘çÜ"
'    End If
'    B2Op_Hinmoku(0).Tag = Format(INDEX, "0")
'
'    For Each gObj In Back_Juchu
'        gObj.Visible = B2Op_Hinmoku(0).Value
'    Next gObj
'    For Each gObj In Back_Hachu
'        gObj.Visible = B2Op_Hinmoku(1).Value
'    Next gObj
'    If FormAct Then
'       Call iB2_Torcd_LostFocus
'    End If
'End Sub
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'---------------------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA end

'+---( B3 )----------------------------------------------------------------------------
'+-----------------------------+
'+  B3Op_SKKbn
'+-----------------------------+
Private Sub B3Op_SKKbn_Click(INDEX As Integer)
    If INDEX = 0 Then
        B3Op_SKKbn(0).BackColor = mGotColor
        B3Op_SKKbn(1).BackColor = mLostColor
        Back_Hachu(8).Enabled = False
        B3Chk_NHKbn.Value = 0
    Else
        B3Op_SKKbn(0).BackColor = mLostColor
        B3Op_SKKbn(1).BackColor = mGotColor
        Back_Hachu(8).Enabled = True
    End If
End Sub
'+-----------------------------+
'+  iB3_Hako1
'+-----------------------------+
Private Sub iB3_Hako1_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_Hako1_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB3_CDate
'+-----------------------------+
Private Sub iB3_CDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_CDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB3_Hako2
'+-----------------------------+
Private Sub iB3_Hako2_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_Hako2_LostFocus()
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
    If TorNmGet(iB3_KTorcd, 5) Then
        B3lb_KTorcd = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  iB3_MTorcd
'+-----------------------------+
Private Sub iB3_MTorcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_MTorcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B3lb_MTorcd = ""
    If TorNmGet(iB3_MTorcd, 8) Then
        B3lb_MTorcd = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  iB3_MUkeir
'+-----------------------------+
Private Sub iB3_MUkeir_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_MUkeir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB3_SDate
'+-----------------------------+
Private Sub iB3_SDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_SDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB3_EDate
'+-----------------------------+
Private Sub iB3_EDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_EDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  B3Comb_THKbn
'+-----------------------------+
Private Sub B3Comb_THKbn_GotFocus()
    B3Comb_THKbn.BackColor = gIGotSel
End Sub
Private Sub B3Comb_THKbn_LostFocus()
    B3Comb_THKbn.BackColor = gILostSel
End Sub
Private Sub B3Comb_THKbn_Click()
    CisComboTH.CodeGet
    gStr = CisComboTH.Code
'======================================== ***** 2008/06/11 Start
'    Select Case gStr
'        Case "1"        ' íºëó
'            Back_Hachu(2) = "î[ì¸êÊ(íºëó)"
'            Back_Hachu(2).Tag = "X"
'        Case Else
'            Back_Hachu(2) = "ÉÅÅ[ÉJÅ["
'            Back_Hachu(2).Tag = ""
'    End Select
    
'-----------------------------------------------------------' 2009.02.27 del start
'    If B2Op_Hinmoku(0).Value Then
'       PB_KTN.Visible = False: PB_STN.Visible = False
'       Exit Sub
'    End If
'-----------------------------------------------------------' 2009.02.27 del end
    
    Select Case gStr
        Case "1"        ' íºëó
            Back_Hachu(2) = "î[ì¸êÊ(íºëó)"
            Back_Hachu(2).Tag = "X"
            '========================================= 2009/01/15 Start
            'Back_Hachu(24).Visible = False
            'Back_Hachu(25).Visible = True
            Back_Hachu(24).Visible = True
            Back_Hachu(25).Visible = False
            '========================================= 2009/01/15 End
            
            Back_Hachu(14).Visible = True: Back_Hachu(15).Visible = True: Back_Hachu(19).Visible = True
            Back_Hachu(16).Visible = True: Back_Hachu(17).Visible = True: Back_Hachu(18).Visible = True
            Back_Hachu(26).Visible = False
            Back_Hachu(25).Visible = False                                      ' 2009.02.27 add
            Back_Hachu(41).Visible = True: Back_Hachu(42).Visible = True        ' 2009.02.27 add
            
            PB_KTN.Visible = True: PB_STN.Visible = True
        Case "4"
            Back_Hachu(2) = "ÉÅÅ[ÉJÅ["
            Back_Hachu(2).Tag = ""
            Back_Hachu(24).Visible = True
            Back_Hachu(25).Visible = False
            
            Back_Hachu(14).Visible = False: Back_Hachu(15).Visible = False: Back_Hachu(19).Visible = False
            Back_Hachu(16).Visible = False: Back_Hachu(17).Visible = False: Back_Hachu(18).Visible = False
            Back_Hachu(26).Visible = True
            Back_Hachu(25).Visible = True                                       ' 2009.02.27 add
            Back_Hachu(41).Visible = False: Back_Hachu(42).Visible = False      ' 2009.02.27 add
            PB_KTN.Visible = False: PB_STN.Visible = True
        
        Case Else
            Back_Hachu(2) = "ÉÅÅ[ÉJÅ["
            Back_Hachu(2).Tag = ""
            Back_Hachu(24).Visible = False
            Back_Hachu(25).Visible = True
            
            Back_Hachu(14).Visible = True: Back_Hachu(15).Visible = True: Back_Hachu(19).Visible = True
            Back_Hachu(16).Visible = True: Back_Hachu(17).Visible = True: Back_Hachu(18).Visible = True
            Back_Hachu(26).Visible = False
            Back_Hachu(25).Visible = False                                      ' 2009.02.27 add
            Back_Hachu(41).Visible = True: Back_Hachu(42).Visible = True        ' 2009.02.27 add
            
            PB_KTN.Visible = True: PB_STN.Visible = False
    End Select
'======================================== ***** 2008/06/11 End
End Sub
'======================================== ***** 2008/06/11 Start
'+-----------------------------+
'+  B3Chk_PrtkbSNÅ@éËîzãÊï™ÅFéxãã/íºëó î[ïièëî≠çsãÊï™
'+-----------------------------+
Private Sub B3Chk_PrtkbSN_Click()
    If B3Chk_PrtkbSN.Value = 0 Then
       B3Chk_PrtkbSN.Caption = "î≠çsóL"
    Else
       B3Chk_PrtkbSN.Caption = "î≠çsñ≥"
    End If
End Sub
'======================================== ***** 2008/06/11 End
'+-----------------------------+
'+  B3Chk_Week
'+-----------------------------+
Private Sub B3Chk_Week_Click(INDEX As Integer)
    If B3Chk_Week(INDEX).Value = 1 Then
       B3Chk_Week(INDEX).BackColor = mGotColor
    Else
       B3Chk_Week(INDEX).BackColor = mLostColor
    End If
End Sub
'+-----------------------------+
'+  B3Chk_NHKbn ( 1/Ç∑ÇÈ)
'+-----------------------------+
Private Sub B3Chk_NHKbn_Click()
    If B3Chk_NHKbn.Value = 0 Then
        B3Chk_NHKbn.BackColor = mLostColor
        B3Chk_NHKbn.Caption = " ÇµÇ»Ç¢"
    Else
        B3Chk_NHKbn.BackColor = mGotColor
        B3Chk_NHKbn.Caption = "Ç∑ÇÈ"
    End If
End Sub
'+-----------------------------+
'+  iB3_Densy
'+-----------------------------+
Private Sub iB3_Densy_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_Densy_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    B3lb_Densy = ""
    If MeisyoGet("ì`ï[éÌóﬁî≠íç", iB3_Densy) Then
        B3lb_Densy = SYM_Meisyo
    End If
End Sub
'''2007/12/11DEL'+----------------------------------+
'''2007/12/11DEL'+  B3Comb_NouhinÅ@î[ïièëå`ë‘
'''2007/12/11DEL'+----------------------------------+
'''Private Sub B3Comb_Nohin_GotFocus()
'''2007/12/11DEL    B3Comb_Nohin.BackColor = gIGotSel
'''End Sub
'''Private Sub B3Comb_Nohin_LostFocus()
'''2007/12/11DEL    B3Comb_Nohin.BackColor = gILostSel
'''End Sub
'''2007/12/11DEL'+-----------------------------+
'''2007/12/11DEL'+  B1Comb_JyuryoÅ@éÛóÃèëå`ë‘
'''2007/12/11DEL'+-----------------------------+
'''Private Sub B3Comb_Jyuro_GotFocus()
'''2007/12/11DEL    B3Comb_Jyuro.BackColor = gIGotSel
'''End Sub
'''Private Sub B3Comb_Jyuro_LostFocus()
'''2007/12/11DEL    B3Comb_Jyuro.BackColor = gILostSel
'''End Sub
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
'+-----------------------------+
'+  B3Chk_JRKbn
'+-----------------------------+
Private Sub B3Chk_JRKbn_Click()
    If B3Chk_JRKbn.Value = 0 Then
        B3Chk_JRKbn.Caption = "î≠íçéû"
    Else
        B3Chk_JRKbn.Caption = "éÛóÃéû"
    End If
End Sub
'+-----------------------------+
'+  B3Comb_MeisiÅ@ñæç◊èëî≠çs
'+-----------------------------+
Private Sub B3Comb_Meisi_GotFocus()
    B3Comb_Meisi.BackColor = gIGotSel
End Sub
Private Sub B3Comb_Meisi_LostFocus()
    B3Comb_Meisi.BackColor = gILostSel
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

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
'+-----------------------------+
'+  iB3_DensyJ
'+-----------------------------+
Private Sub iB3_DensyJ_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_DensyJ_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    B3lb_DensyJ = ""
    If MeisyoGet("ì`ï[éÌóﬁéÛíç", iB3_DensyJ) Then
        B3lb_DensyJ = SYM_Meisyo
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
'+  B3Chk_JRKbnJ
'+-----------------------------+
Private Sub B3Chk_JRKbnJ_Click()
    If B3Chk_JRKbnJ.Value = 0 Then
        B3Chk_JRKbnJ.Caption = "î≠íçéû"
    Else
        B3Chk_JRKbnJ.Caption = "éÛóÃéû"
    End If
End Sub
'+----------------------+
'+     B3Chk_MaeHak     +
'+----------------------+
Private Sub B3Chk_Maehak_Click()
    If B3Chk_Maehak.Value = 0 Then
       B3Chk_Maehak.Caption = "ï\é¶ñ≥"
    Else
       B3Chk_Maehak.Caption = "ï\é¶óL"
    End If
End Sub
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END

'-------------------------------------------------------------------------------' 2009.09.23 add start
'+-----------------------------+
'+  B3Chk_JIS
'+-----------------------------+
Private Sub B3Chk_JIS_Click()
    If B3Chk_JIS.Value = 0 Then
        B3Chk_JIS.Caption = "ïWèÄ"
    Else
        B3Chk_JIS.Caption = "JISãKäi"
    End If
End Sub
'-------------------------------------------------------------------------------' 2009.09.23 add end
'-------------------------------------------------------------------------------' 2010.01.16 add start
'+-----------------------------+
'+  B3Chk_Tana
'+-----------------------------+
Private Sub B3Chk_Tana_Click()
    If B3Chk_Tana.Value = 0 Then
        B3Chk_Tana.Caption = "Ç∑ÇÈ"
    Else
        B3Chk_Tana.Caption = "ÇµÇ»Ç¢"
    End If
End Sub
'-------------------------------------------------------------------------------' 2009.01.16 add end

'-------------------------------------------------------------------------------' 2009.02.26 add start
''+------------------+
''+    îwåiêF        +
''+------------------+
'Private Sub B3Btn_Hikei_Click()                            2009/12/08
'    With CommonDialog1
'        .CancelError = True
'        On Error GoTo ErrHandler
'
'        .Color = B3Btn_Hikei.BackColor
'        .Flags = cdlCCRGBInit Or cdlCCPreventFullOpen
'        .Flags = cdlCCRGBInit
'        .ShowColor
'
'        B3Btn_Hikei.BackColor = .Color
'    End With
'ErrHandler:
'    On Error GoTo 0
'End Sub

'+-----------------------------+    ' â¡çHêÊ
'+  iB3_STorcd
'+-----------------------------+
Private Sub iB3_STorcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_STorcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B3lb_STorcd = ""
    If TorNmGet(iB3_STorcd, 7) Then     ' 2009.04.07 upd
        B3lb_STorcd = TRM_RName
    End If
End Sub
'+-----------------------------+    ' ìæà”êÊ
'+  iB3_TTorcd
'+-----------------------------+
Private Sub iB3_TTorcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_TTorcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B3lb_TTorcd = ""
    If TorNmGet(iB3_TTorcd, 2) Then
        B3lb_TTorcd = TRM_RName
    End If
End Sub
'+----------------------+
'+     B3Chk_KanBn      +
'+----------------------+
Private Sub B3Chk_KanBn_Click()
    If B3Chk_KanBn.Value = 0 Then
       B3Chk_KanBn.Caption = "ëÂ"
    Else
       B3Chk_KanBn.Caption = "è¨"
    End If
End Sub
'+----------------------+
'+     B3Chk_GenPn      +
'+----------------------+
Private Sub B3Chk_GenPn_Click()
    If B3Chk_GenPn.Value = 0 Then
       B3Chk_GenPn.Caption = "ëÂ"
    Else
       B3Chk_GenPn.Caption = "è¨"
    End If
End Sub
'-------------------------------------------------------------------------------' 2009.02.26 add end
'======================================== 2009/12/08 Start
'+-----------------------------+
'+  iB3_Okiba1
'+-----------------------------+
Private Sub iB3_Okiba1_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_Okiba1_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
'    B3lb_XXXXX = ""
'    If MeisyoGet("ä«óùèÍèä", iB3_Okiba1) Then
'        B3lb_XXXXX = SYM_Meisyo
'    End If
End Sub
'+-----------------------------+
'+  iB3_Okiba2
'+-----------------------------+
Private Sub iB3_Okiba2_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB3_Okiba2_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
'    B3lb_XXXXX = ""
'    If MeisyoGet("ä«óùèÍèäÉÜÅ[ÉUÅ[", iB3_Okiba2) Then
'        B3lb_XXXXX = SYM_Meisyo
'    End If
End Sub
'======================================== 2009/12/08 End

'**************************************
'*                                    *
'*      èàóùãÊï™Ç…ÇÊÇËãÊï™ñºï\é¶      *
'*                                    *
'**************************************
'+-----------------------+
'+      äÓñ{èÓïÒ
'+-----------------------+
Private Sub SyoriName(SYKB As String)
   Dim W_Color     As Long
   W_Color = &H80000008 'èâä˙êFæØƒ(çï)
   
   H2LB_Size_B.Enabled = False
   
   Select Case SyoriKB
     Case "ADD"
          SyoriNM = "í«â¡"
          K_Sykbnm.ForeColor = &HFFFF&   '(â©êF)
          PB_ADD.Visible = False
              
          H2LB_Size_B.Enabled = True
     Case "MNT"
          SyoriNM = "èCê≥"
          K_Sykbnm.ForeColor = &HFF0000  '(ê¬êF)
'*...............................................:: 2004.06.21 :: -- Insert Start
          H2LB_Size_B.Enabled = True
'*^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^:: 2004.06.21 :: -- Insert End
     
     Case "DEL"
          SyoriNM = "çÌèú"
          K_Sykbnm.ForeColor = &HFF&     '(ê‘êF)
          W_Color = &HFF&
'*...............................................:: 2004.06.21 :: -- Insert Start
          H2LB_Size_B.Enabled = True
'*^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^:: 2004.06.21 :: -- Insert End
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
    Dim H2Mode      As Boolean
    Dim H2Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long
    Dim B2Mode      As Boolean
    Dim B2Color     As Long
    Dim B3Mode      As Boolean
    Dim B3Color     As Long
    Dim B4Mode      As Boolean
    Dim B4Color     As Long
    
    H1Mode = False
    H1Color = gPLostSel
    H2Mode = False
    H2Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
    B2Mode = False
    B2Color = gPLostSel
    B3Mode = False
    B3Color = gPLostSel
    B4Mode = False
    B4Color = gPLostSel

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
       Case "H2"
          H2Mode = True
          H2Color = gPGotSel
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
          PB_CHG.Tag = "TOR"
          If CisFun.Val2(PB_Hin.Tag) = 0 Then
            PB_CHG.BtnCaption = "éËîzêÊ"
          Else
            PB_CHG.BtnCaption = "ïiî‘"
          End If
       Case "B2"
          B2Mode = True
          B2Color = gPGotSel
          PB_CHG.Tag = "BASE"
          PB_CHG.BtnCaption = "äÓñ{"
       Case "B3"
          B3Mode = True
          B3Color = gPGotSel
       Case "B4"
          B4Mode = True
          B4Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    H2_Area1.Enabled = H2Mode
    H2_Area1.BackColor = H2Color
    
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B1_Area2.Enabled = B1Mode Or B2Mode
    B1_Area2.BackColor = B1Color
    If B2Mode Then
        VsGridT.TabStop = False
    Else
        VsGridT.TabStop = True
    End If
    
    B2_Area1.Enabled = B2Mode
    B2_Area1.BackColor = B2Color
    
    B3_Area1.Enabled = B3Mode
    B3_Area1.BackColor = B3Color
'    B4_Area1.Enabled = B1Mode Or B4Mode    ' 2009.03.02 del
    B4_Area1.Enabled = B3Mode Or B4Mode     ' 2009.03.02 add
    B4_Area1.BackColor = B4Color
    
    PB_Look.Visible = H1Mode
    PB_ADD.Visible = Not B1Mode And Not B3Mode
    PB_MNT.Visible = Not B1Mode And Not B3Mode
    PB_DEL.Visible = Not B1Mode And Not B3Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    If Rv_Call3 <> "" Then
        PB_CAN.Visible = Not B2Mode
    End If
    
    PB_Hin.Visible = B1Mode
    
    If Rv_Call3 = "" Then
        PB_CHG.Visible = B1Mode Or B2Mode
    Else
        PB_CHG.Visible = False
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
    B3_Area3.Visible = B3Mode
'    PB_KTN.Visible = B3Mode And B2Op_Hinmoku(1).Value
'    PB_STN.Visible = B3Mode
'    PB_KTN.Enabled = B3Mode And B2Op_Hinmoku(1).Value
'    PB_STN.Enabled = B3Mode
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END

'---------------------------------------------------------------' 2009.09.23 add start
    PB_JIS.Visible = B1Mode                 ' JISÉ{É^Éì
    B5_JISNm.Visible = Not H1Mode           ' çÄñ⁄ñº
    B5_JISNo.Visible = Not H1Mode           ' JISãKäi
'---------------------------------------------------------------' 2009.09.23 add end

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
            If .Name Like wbc_Enm & "Btn_*" Then .BackColor = mBackColor    ' 2009.02.27 add
        End With
    Next gObj
    
    
    If HeadBody = "H2" Or HeadBody = "H" Or HeadBody = "ALL" Then
        CisVsGrid.Clear
        Call CisVsGridT.Clear
    End If
    If HeadBody = "B1" Or HeadBody = "B" Or HeadBody = "ALL" Then
        
        CisVsGridT.Clear
'        Call GridInitT
        CisVsGridH.Clear
'        Call GridInitH
    End If
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
    
' H1
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        If H1Chk_Size.Value = 0 Then
            ProcHB = "H2"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
        
            If SyoriKB = "ADD" Then
                iH2_Itatu.SetFocus
            Else
               vsGrid1.SetFocus
               vsGrid1.Row = 1
               Call VSGrid1_ClickClick
            End If
        Else
            If SyoriKB = "DEL" Then
                If Not DBDelete Then GoTo ReturnPress_Ed
                Call HeadBodyClear("B")
                ProcHB = "H1"
                Call DispChange(ProcHB)
                Call SyoriName(SyoriKB)
                GoTo ReturnPress_Ed
            End If
            
            Call GridSetT
            Call GridSetH               ' 2009.03.03 add
            ProcHB = "B1"
            Call DispChange(ProcHB)
            iB1_Name.SetFocus
            
        End If
        GoTo ReturnPress_Ed
    End If

' H2
    If ProcHB = "H2" Then
        If Not Head2Chk Then GoTo ReturnPress_Ed
' << çÌèúèàóù >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call HeadBodyClear("H2")
            Call Head2Input("INV")
            
            If RTrim(vsGrid1.TextMatrix(1, 4)) = "" Then
                Call SyoriName(SyoriKB)
                ProcHB = "H1"
                Call DispChange(ProcHB)
            Else
                Call VSGrid1_ClickClick
            End If
            GoTo ReturnPress_Ed
        Else
            ProcHB = "B1"
            Call DispChange(ProcHB)
'            If SyoriKB = "ADD" Then: Call CisComboZK.ComboSet
            If SyoriKB = "ADD" Then
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/05/15 INSERT START
                wKanri = 0
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/05/15 INSERT END
                Call CisComboZK.ComboSet
                iB1_Name.SetFocus
                B1Op_Syoyo(0).Value = True      ' 2009.02.26 add
            Else
                VsGridT.SetFocus
            End If
            iB1_Name.SetFocus
            GoTo ReturnPress_Ed
        End If
        
    End If

' B1
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut(True) Then
            GoTo ReturnPress_Ed
        End If
        Call HeadBodyClear("B")
        Call HeadBodyClear("H2")
        CisVsGrid.Clear
        RowColFlg = True        ' 2008/02/06 ADD
        Call CisVsGridT.Clear
        RowColFlg = False       ' 2008/02/06 ADD
'        Call CisVsGridH.Clear
        
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        GoTo ReturnPress_Ed
    End If

' B2
    If ProcHB = "B2" Then
        If Not Body2Chk Then GoTo ReturnPress_Ed
' << çÌèúèàóù >>
        If SyoriKB2 = "DEL" Then
            If Not DBDeleteTor Then
               RowColFlg = True
               Call CisVsGridT.vsColor(è¡ãéëSïî)
               RowColFlg = False
               GoTo ReturnPress_Ed
            End If
            Call HeadBodyClear("B3")
            Call DispChange(ProcHB)
            Call SyoriName2(SyoriKB2)
        Else
            GridBack = ProcHB
            ProcHB = "B3"
            Call DispChange(ProcHB)
            Call vsGridH_Select                 ' 2009.03.03 add
        End If
        GoTo ReturnPress_Ed
    End If

' B3
    If ProcHB = "B3" Then
        If Not Body3Chk Then GoTo ReturnPress_Ed

        RowColFlg = True
        Call CisVsGridT.vsColor(è¡ãéëSïî)
        RowColFlg = False

        If Not DBPutTor Then GoTo ReturnPress_Ed

'-----------------------------------------------' 2009.03.03 add start
        RowColFlg = True
        Call CisVsGridH.vsColor(è¡ãéëSïî)
        RowColFlg = False
'-----------------------------------------------' 2009.03.03 add end
        
        Call HeadBodyClear("B3")
'        Call Head2Input("INV")
        Call GridSetT
        
        ProcHB = "B2"
        Call DispChange(ProcHB)
        iB2_Torcd.SetFocus
        If GridBack = "B1" Then
           SyoriKB2 = ""
           Call SyoriName2(SyoriKB2)
           Call DispChange(ProcHB)
'           VsGrid1.SetFocus
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
' çﬁéø
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

    Call HeadBodyClear("H2")
' *DEL* 2004/04/12
'    Call VsGridT.Clear
    wKanri = 0
' É}ÉXÉ^ì¸óÕ SELECTï∂ æØƒ
    gSL_Select = " SELECT * "
    gSL_Select = gSL_Select & " FROM çﬁóøÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " WHERE çﬁéø = '" & RTrim(iH1_Zaist) & "'"
    If Not ZRMRead(gSL_Select, 1) Then
       If SyoriKB <> "ADD" Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    çﬁóøÉ}ÉXÉ^ñ¢ìoò^            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          iH1_Zaist.SetFocus
          Exit Function
       End If
    Else
        If H1Chk_Size.Value <> ZRM.ê°ñ@ãÊï™ Then: H1Chk_Size.Value = ZRM.ê°ñ@ãÊï™
    End If

    If Head2Input("INV") Then
        If Rv_Call3 = "" Then
            If H1Chk_Size.Value <> 0 Then
                If SyoriKB = "ADD" Then
                   With CisFun
                       .MB_Lines = 4
                       .MB_MSG(2) = "    çﬁóøÉ}ÉXÉ^ìoò^çœ            "
                       .MB_Title = ""
                       .MB_Button = Error
                       .MBOX
                   End With
                   iH1_Zaist.SetFocus
                   Exit Function
                End If
            End If
        End If
    Else
        If Rv_Call3 = "" Then
            If H1Chk_Size.Value <> 0 Then
                If SyoriKB <> "ADD" Then
                    With CisFun
                        .MB_Lines = 4
                        .MB_MSG(2) = "    çﬁóøÉ}ÉXÉ^ñ¢ìoò^            "
                        .MB_Title = ""
                        .MB_Button = Error
                        .MBOX
                    End With
                    iH1_Zaist.SetFocus
                    Exit Function
                    
                    wKanri = 0
                End If
            End If
        End If
    End If
    
'---------------------------------------------------------------------------------------' 2009.09.23-24 add start
    ' JISãKäiì¸óÕ
    B5_JISNo.Caption = ""
    gSL_Select = " SELECT * FROM çﬁóøJISãKäiïœä∑É}ÉXÉ^ "
    gSL_Select = gSL_Select & " WHERE çﬁéø = '" & RTrim(iH1_Zaist) & "'"
    If ZJSRead(gSL_Select, 1) Then
        B5_JISNo.Caption = ZJS.JISãKäi
    End If
'---------------------------------------------------------------------------------------' 2009.09.23-24 add end
    
    Head1Chk = True
End Function
'****************************
'*      Head2¡™Ø∏èàóù
'****************************
Private Function Head2Chk() As Boolean
    Head2Chk = False

'*...............................................:: 2004.06.21 :: -- Delete Start
'   If SyoriKB = "ADD" Then
'*^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^:: 2004.06.21 :: -- Delete End
        If CisFun.Val2(iH2_Itatu) = 0 And _
           CisFun.Val2(iH2_Width) = 0 And _
           CisFun.Val2(iH2_Long) = 0 Then
            With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    ê°ñ@èÓïÒÇ™ì¸óÕÇ≥ÇÍÇƒÇ¢Ç‹ÇπÇÒ         "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
            End With
            iH2_Itatu.SetFocus
            Exit Function
        End If
'*...............................................:: 2004.06.21 :: -- Delete Start
'   Else
'       If CisFun.Val2(VsGrid1.TextMatrix(VsGrid1.Row, 4)) = 0 Then
'           With CisFun
'              .MB_Lines = 4
'              .MB_MSG(2) = "    ê°ñ@èÓïÒÇ™ëIëÇ≥ÇÍÇƒÇ¢Ç‹ÇπÇÒ        "
'              .MB_Title = ""
'              .MB_Button = Error
'              .MBOX
'           End With
'           VsGrid1.SetFocus
'           Exit Function
'       End If
'   End If
'*^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^:: 2004.06.21 :: -- Delete End

    gSL_Select = "SELECT * FROM çﬁóøÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " WHERE çﬁéø = '" & RTrim(iH1_Zaist) & "'"
    gSL_Select = gSL_Select & "   AND î¬å˙ =  " & CisFun.Val2(iH2_Itatu)
    gSL_Select = gSL_Select & "   AND ïù =    " & CisFun.Val2(iH2_Width)
    gSL_Select = gSL_Select & "   AND í∑Ç≥ =  " & CisFun.Val2(iH2_Long)
    If ZRMRead(gSL_Select, 1) Then
        If SyoriKB = "ADD" Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    çﬁóøÉ}ÉXÉ^ìoò^çœ            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iH2_Itatu.SetFocus
           Exit Function
        End If
    Else
        If SyoriKB <> "ADD" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    çﬁóøÉ}ÉXÉ^ñ¢ìoò^            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            vsGrid1.SetFocus
            Exit Function
            
            wKanri = 0
        End If
    End If

'*...............................................:: 2004.06.21 :: -- Insert Start
'   >> çsà⁄ìÆÇ…ÇÊÇÈëIëÅïï\é¶
    If SyoriKB <> "ADD" Then
        With vsGrid1
            For mIx = 1 To .Rows - 1
                If CisFun.Val2(.TextMatrix(mIx, SCol(1))) = CisFun.Val2(iH2_Itatu) And _
                   CisFun.Val2(.TextMatrix(mIx, SCol(2))) = CisFun.Val2(iH2_Width) And _
                   CisFun.Val2(.TextMatrix(mIx, SCol(3))) = CisFun.Val2(iH2_Long) Then
                    If .Row <> mIx Then
                        .Row = mIx
                        vsGrid1.SetFocus
                        SendKeys "{LEFT}"
                        DoEvents
                    End If
                    Exit For
                End If
            Next mIx
        End With
    End If
'*^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^:: 2004.06.21 :: -- Insert End

    Head2Chk = True
End Function
'****************************
'*      Body1¡™Ø∏èàóù
'****************************
Private Function Body1Chk() As Boolean
    
    Body1Chk = False
    ' çﬁóøñº
'============ 2010/03/18    Start
'    If Trim(iB1_Name) = "" Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    çﬁóøñºÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
'            .MB_Title = ""
'            .MB_Button = Error
'            .MBOX
'        End With
'       iB1_Name.SetFocus
'       Exit Function
'    End If
'============ 2010/03/18    End
    ' çﬁéÌ
    If RTrim(iB1_Zaisy) <> "" Then
        If Not MeisyoGet("çﬁéÌãÊï™", iB1_Zaisy) Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    çﬁéÌãÊï™  ñ¢ìoò^          "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
            iB1_Zaisy.SetFocus
            Exit Function
        End If
    End If
    ' å`èÛ
    If RTrim(iB1_Keijo) <> "" Then
        If Not MeisyoGet("å`èÛãÊï™", iB1_Keijo) Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    å`èÛãÊï™  ñ¢ìoò^          "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
            iB1_Keijo.SetFocus
            Exit Function
        End If
    End If
        
    ' égópíPà 
    If RTrim(iB1_Tani) = "" And _
       RTrim(iB1_TTani) <> "" Then
       iB1_Tani = iB1_TTani
    End If
        
    If RTrim(iB1_Tani) <> "" Then
        If Not MeisyoGet("íPà ãÊï™", iB1_Tani) Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    égópíPà   ñ¢ìoò^          "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
            iB1_Tani.SetFocus
            Exit Function
        End If
    End If
        
    ' éËîzíPà 
'============ 2010/03/18    Start
'    If RTrim(iB1_TTani) = "" And _
'       RTrim(iB1_Tani) <> "" Then
'       iB1_TTani = iB1_Tani
'    End If
    If RTrim(iB1_TTani) = "" Then
       iB1_TTani = "22"
       Call iB1_TTani_LostFocus
    End If
'============ 2010/03/18    End
    
    If RTrim(iB1_TTani) = "" Then
         With CisFun
             .MB_Lines = 4
             .MB_MSG(2) = "    éËîzíPà Çì¸óÕÇµÇƒâ∫Ç≥Ç¢          "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
         End With
        iB1_TTani.SetFocus
        Exit Function
    Else
        If Not MeisyoGet("íPà ãÊï™", iB1_TTani) Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    éËîzíPà   ñ¢ìoò^          "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
            iB1_TTani.SetFocus
            Exit Function
        End If
    End If
        
    ' ç›å…ä«óù
    With CisComboZK
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
           Exit Function
        End If
    End If
    
'====================================================== *** 2008/06/27 Start(Comment)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
'    If RTrim(iB1_JHinbn) <> "" Then
'       gSL_Select = ""
'       gSL_Select = gSL_Select & " SELECT * FROM çﬁóøÉ}ÉXÉ^"
'       gSL_Select = gSL_Select & "  WHERE éÛããïiî‘ = '" & RTrim(iB1_JHinbn) & "'"
'       If SyoriKB = "MNT" Then
'          gSL_Select = gSL_Select & "   AND çﬁóøä«óùî‘çÜ <> " & wKanri
'       End If
'       If ZRMRead(gSL_Select, 1) Then
'          With CisFun
'             .MB_Lines = 5
'             .MB_MSG(2) = "    éwíËÇ≥ÇÍÇΩéÛããïiî‘ÇÕ     "
'             .MB_MSG(4) = "     ä˘Ç…ìoò^çœÇ›Ç≈Ç∑ÅB      "
'             .MB_Title = ""
'             .MB_Button = Error
'             .MBOX
'          End With
'          iB1_JHinbn.SetFocus
'          Exit Function
'       End If
'    End If
'
'    If RTrim(iB1_HHinbn) <> "" Then
'       gSL_Select = ""
'       gSL_Select = gSL_Select & " SELECT * FROM çﬁóøÉ}ÉXÉ^"
'       gSL_Select = gSL_Select & "  WHERE î≠íçïiî‘ = '" & RTrim(iB1_HHinbn) & "'"
'       If SyoriKB = "MNT" Then
'          gSL_Select = gSL_Select & "   AND çﬁóøä«óùî‘çÜ <> " & wKanri
'       End If
'       If ZRMRead(gSL_Select, 1) Then
'          With CisFun
'             .MB_Lines = 5
'             .MB_MSG(2) = "    éwíËÇ≥ÇÍÇΩî≠íçïiî‘ÇÕ     "
'             .MB_MSG(4) = "     ä˘Ç…ìoò^çœÇ›Ç≈Ç∑ÅB      "
'             .MB_Title = ""
'             .MB_Button = Error
'             .MBOX
'          End With
'          iB1_HHinbn.SetFocus
'          Exit Function
'       End If
'    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'====================================================== *** 2008/06/27 End(Comment)
    
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
            .MB_MSG(2) = "    édì¸êÊÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
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
'---------------------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA start
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE START
'    'gSL_Select = gSL_Select & " WHERE éÊà¯êÊãÊï™ = 1 "
'    gSL_Select = gSL_Select & " WHERE éÊà¯êÊãÊï™ = " & B2Op_Hinmoku(0).Tag
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE END
'---------------------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA end
    gSL_Select = gSL_Select & " WHERE éÊà¯êÊãÊï™ = 1 "                                  ' 2009.02.27 add : ïiñ⁄ä÷òA
    gSL_Select = gSL_Select & "   AND éÊà¯êÊCD = '" & RTrim(iB2_Torcd) & "'"
    If Not TRMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     édì¸êÊ ñ¢ìoò^              "
            .MB_MSG(3) = "    ÅyéÊà¯êÊÉ}ÉXÉ^Åz            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB2_Torcd.SetFocus
       Exit Function
    End If
    
' éÛì¸ë∂ç›É`ÉFÉbÉN
    If Trim(iB2_Ukeir) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    éÛì¸Çì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB2_Ukeir.SetFocus
       Exit Function
    End If
    
    If Trim(iB2_Ukeir) <> "" Then
        gSL_Select = " SELECT * "
        gSL_Select = gSL_Select & " FROM éÊà¯êÊéÛì¸É}ÉXÉ^ "
'---------------------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA start
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE START
'        'gSL_Select = gSL_Select & " WHERE éÊà¯êÊãÊï™ = 1 "
'        gSL_Select = gSL_Select & " WHERE éÊà¯êÊãÊï™ = " & B2Op_Hinmoku(0).Tag
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE END
'---------------------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA end
        gSL_Select = gSL_Select & " WHERE éÊà¯êÊãÊï™ = 1 "                              ' 2009.02.27 add : ïiñ⁄ä÷òA
        gSL_Select = gSL_Select & "   AND éÊà¯êÊCD = '" & RTrim(iB2_Torcd) & "'"
        gSL_Select = gSL_Select & "   And éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
        If Not TUMRead(gSL_Select, 1) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    édì¸êÊÇ…éÛì¸Ç™ìoò^Ç≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÅB            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iB2_Ukeir.SetFocus
           Exit Function
        End If
    End If
' É}ÉXÉ^ì¸óÕ SELECTï∂ æØƒ
    With VsGridT
        ReadFlg = False
        B2Op_Hinmoku(1).Value = True                                                ' 2009.02.27 add : ïiñ⁄ä÷òA
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
        If B2Op_Hinmoku(0).Value Then
           gStr = "0"
        Else
           gStr = "1"
        End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
        iB2_Torcd.Tag = ""
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
        For mIx = 1 To .Rows - 1
            If Trim(.TextMatrix(mIx, SColT(35))) = "" Then Exit For
            If RTrim(iB2_Torcd) = Trim(.TextMatrix(mIx, SColT(35))) And _
               RTrim(iB2_Ukeir) = Trim(.TextMatrix(mIx, SColT(2))) Then
               If RTrim(gStr) <> Trim(.TextMatrix(mIx, SColT(39))) Then
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
            If RTrim(iB2_Torcd) = Trim(.TextMatrix(mIx, SColT(35))) And _
               RTrim(iB2_Ukeir) = Trim(.TextMatrix(mIx, SColT(2))) And _
               RTrim(gStr) = Trim(.TextMatrix(mIx, SColT(39))) Then
               ReadFlg = True
               iB2_Torcd.Tag = mIx
               .Row = mIx
               Exit For
            End If
        Next mIx
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
        For mIx = 1 To .Rows - 1
            If Trim(.TextMatrix(mIx, SColT(35))) = "" Then Exit For
            If RTrim(iB2_Torcd) = Trim(.TextMatrix(mIx, SColT(35))) And _
               RTrim(iB2_Ukeir) = Trim(.TextMatrix(mIx, SColT(2))) Then
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
                   .MB_MSG(2) = "    çﬁóøéÊà¯êÊÉ}ÉXÉ^ìoò^çœ            "
                   .MB_Title = ""
                   .MB_Button = Error
                   .MBOX
              End With
              iB2_Torcd.SetFocus
              Exit Function
           End If
           ' éÊà¯êÊèÓïÒï\é¶
           Call HeadBodyClear("B3")
           RowColFlg = True
'           Call CisVsGrid.vsColor(å¬ï )
           RowColFlg = False

           Call GridToTextT
           DoEvents
       Else
           If SyoriKB2 <> "ADD" Then
              With CisFun
                   .MB_Lines = 4
                   .MB_MSG(2) = "    çﬁóøéÊà¯êÊÉ}ÉXÉ^ñ¢ìoò^            "
                   .MB_Title = ""
                   .MB_Button = Error
                   .MBOX
              End With
              iB2_Torcd.SetFocus
              Exit Function
           End If
       
            B3Op_SKKbn(0).Value = True
            Call CisComboTH.ComboSet
'''2007/12/11DEL            Call CisComboNH.ComboSet
'''2007/12/11DEL            Call CisComboJR.ComboSet
            Call CisComboMS.ComboSet
       
            iB3_KTorcd = TRM.å_ñÒêÊ
            If Trim(TUM.ÉTÉCÉNÉã) <> "" Then
               iB3_Cycle1 = Mid(TUM.ÉTÉCÉNÉã, 1, 1)
               iB3_Cycle2 = Mid(TUM.ÉTÉCÉNÉã, 2, 2)
               iB3_Cycle3 = Mid(TUM.ÉTÉCÉNÉã, 4, 2)
            End If
            B3Chk_NHKbn.Value = TRM.î≠íçì‡é¶
'''2007/12/11DEL             ' î[ïièëå`ë‘
'''2007/12/11DEL             With CisComboNH
'''2007/12/11DEL                  .Code = TRM.î[ïièëå`ë‘
'''2007/12/11DEL                  .TextGet
'''2007/12/11DEL             End With
'''2007/12/11DEL             ' éÛóÃèëå`ë‘
'''2007/12/11DEL             With CisComboJR
'''2007/12/11DEL                  .Code = TRM.éÛóÃèëå`ë‘
'''2007/12/11DEL                  .TextGet
'''2007/12/11DEL             End With
             ' ñæç◊èëî≠çs
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE START
'             With CisComboMS
'                  .Code = TRM.ñæç◊èëî≠çs
'                  .TextGet
'             End With
'             B3Chk_PrtkbN.Value = TRM.î[ïièëî≠çsãÊï™
'             B3Chk_PrtkbS.Value = TRM.éxããèëî≠çsãÊï™
'             B3Chk_PrtkbJ.Value = TRM.éÛóÃèëî≠çsãÊï™
'
'             B3Chk_JRKbn.Value = TRM.éÛóÃèëî≠çs
'             Call iB3_KTorcd_LostFocus
                   
           B2Op_Hinmoku(1).Value = True                 ' 2009.02.27 add : ïiñ⁄ä÷òA
           If B2Op_Hinmoku(0).Value Then
              iB3_DensyJ = TRM.ì`ï[éÌóﬁ
              Call iB3_DensyJ_LostFocus
              B3Chk_PrtkbNJ.Value = TRM.î[ïièëî≠çsãÊï™
              B3Chk_PrtkbSJ.Value = TRM.éxããèëî≠çsãÊï™
              B3Chk_PrtkbJJ.Value = TRM.éÛóÃèëî≠çsãÊï™

              B3Chk_JRKbnJ.Value = TRM.éÛóÃèëî≠çs
           Else
              iB3_Densy = TRM.ì`ï[éÌóﬁ
              Call iB3_Densy_LostFocus
              With CisComboMS
                  .Code = TRM.ñæç◊èëî≠çs
                  .TextGet
              End With
              B3Chk_PrtkbN.Value = TRM.î[ïièëî≠çsãÊï™
              B3Chk_PrtkbS.Value = TRM.éxããèëî≠çsãÊï™
              B3Chk_PrtkbJ.Value = TRM.éÛóÃèëî≠çsãÊï™

              B3Chk_JRKbn.Value = TRM.éÛóÃèëî≠çs
              B3Chk_Maehak.Value = TRM.ëOçHíˆãÊï™
           End If
           Call iB3_KTorcd_LostFocus
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE END
       End If
    End With
    
    Body2Chk = True
End Function
'**************************************
'*      BODY3¡™Ø∏èàóù(éÊà¯êÊèÓïÒ)
'**************************************
Private Function Body3Chk() As Boolean
    Dim ReadFlg         As Boolean
    Dim wsTHKbn         As String
    Body3Chk = False

    Call CisVsGridT.EditEnable
    
    '=========================================== *** 2008/06/11 Start
    gStr = CisComboTH.Code
    If RTrim(gStr) = "4" Then
       iB3_Densy = "": B3Chk_PrtkbN.Value = 0: B3Chk_PrtkbS.Value = 0: B3Chk_PrtkbJ.Value = 0: B3Chk_JRKbn.Value = 0
    End If
    '=========================================== *** 2008/06/11 End
    
    
    ' Ç©ÇÒÇŒÇÒéûÅAïKê{ì¸óÕ
    If B3Op_SKKbn(0).Value Then
       ' îwî‘çÜ
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/06/23 DELETE START
'       If Trim(iB3_Seban) = "" Then
'           With CisFun
'               .MB_Lines = 4
'               .MB_MSG(2) = "    îwî‘çÜÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
'               .MB_Title = ""
'               .MB_Button = Error
'               .MBOX
'           End With
'          iB3_Seban.SetFocus
'          Exit Function
'       End If
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/06/23 DELETE END
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
    If Trim(iB3_CDate) <> "" Then
       If Not iB3_CDate.cChkResult Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    ïœçXì˙Çê≥ÇµÇ≠ì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iB3_CDate.SetFocus
          Exit Function
       End If
       ' ïœçXï÷
       If CisFun.Val2(iB3_CBin) = 0 Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    ïœçXï÷Çì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iB3_CBin.SetFocus
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
    ' ïœçXì˙
    If CisFun.Val2(iB3_Syuyo2) <> 0 And Trim(iB3_CDate) = "" Then
       With CisFun
           .MB_Lines = 4
           .MB_MSG(2) = "    ïœçXì˙Çì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
           .MB_Title = ""
           .MB_Button = Error
           .MBOX
       End With
       iB3_CDate.SetFocus
       Exit Function
    End If
    ' å_ñÒêÊ
    If Trim(iB3_KTorcd) = "" Then
        iB3_KTorcd = Trim(iB2_Torcd)
    End If
    B3lb_KTorcd = ""
    B2Op_Hinmoku(1).Value = True                                                    ' 2009.02.27 add : ïiñ⁄ä÷òA
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE START
    If B2Op_Hinmoku(0).Value Then gInt = 2 Else gInt = 5
    'If TorNmGet(iB3_KTorcd, 5) Then
    If TorNmGet(iB3_KTorcd, gInt) Then
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE END
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
'---------------------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA start
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
'    If Not B2Op_Hinmoku(0).Value Then
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'---------------------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA end
'---------------------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA start
'-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´ 2008/02/06 INSERT START
'    End If
'
'    ' ÉAÉhÉåÉX
'    If B2Op_Hinmoku(0).Value Then
'       '================== ïiñ⁄:éÛíç =====================
'       If Trim(iB3_ADD) <> "" Then
'          gSL_Select = "Select * from ïiî‘çﬁóøÉrÉÖÅ[ "
'          gSL_Select = gSL_Select & " Where ÉAÉhÉåÉX = '" & iB3_ADD & "'"
'          gSL_Select = gSL_Select & "   AND ïiñ⁄ = 0"
'          If ZTMRead(gSL_Select, 1) Then
'             If ZTM.ãÊï™ = 0 Then
'                With CisFun
'                     .MB_Lines = 6
'                     .MB_MSG(2) = "    ìØàÍÇÃÉAÉhÉåÉXÇ™ë∂ç›ÇµÇ‹Ç∑ÅB               "
'                     .MB_MSG(4) = "    ïi î‘  ÅF" & ZTM.ï\é¶ïiî‘
'                     .MB_MSG(5) = "    éËîzêÊÅF" & ZTM.éÊà¯êÊ
'                     .MB_MSG(6) = "    éÛ ì¸  ÅF" & ZTM.éÛì¸
'                     .MB_Title = ""
'                     .MB_Button = Error
'                     .MBOX
'                     iB3_ADD.SetFocus
'                     Exit Function
'                End With
'             Else
'                If ZTM.çﬁóøä«óùî‘çÜ <> wKanri Or _
'                   ZTM.éÊà¯êÊ <> iB2_Torcd Or _
'                   ZTM.éÛì¸ <> iB2_Ukeir Then
'                   With CisFun
'                        .MB_Lines = 8
'                        .MB_MSG(2) = "    ìØàÍÇÃÉAÉhÉåÉXÇ™ë∂ç›ÇµÇ‹Ç∑ÅB               "
'                        .MB_MSG(4) = "    çﬁ éø ÅF" & ZTM.çﬁéø
'                        If ZTM.ê°ñ@ãÊï™ = 0 Then
'                           .MB_MSG(5) = "    ê° ñ@ ÅF" & Format(ZTM.î¬å˙, iH2_Itatu.cFormat) & " x " & _
'                                                         Format(ZTM.ïù, iH2_Width.cFormat) & " x " & _
'                                                         Format(ZTM.í∑Ç≥, iH2_Long.cFormat)
'                        End If
'                        .MB_MSG(6) = "    éËîzêÊ ÅF" & ZTM.éÊà¯êÊ
'                        .MB_MSG(7) = "    éÛ ì¸  ÅF" & ZTM.éÛì¸
'                        .MB_Title = ""
'                        .MB_Button = Error
'                        .MBOX
'                   End With
'                   iB3_ADD.SetFocus
'                   Exit Function
'                End If
'             End If
'          End If
'       End If
'    Else
'       '==================== ïiñ⁄:î≠íç ======================
''-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™ 2008/02/06 INSERT END
'---------------------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA end
    ' ÉAÉhÉåÉX
       '==================== ïiñ⁄:î≠íç ======================
       If Trim(iB3_ADD) <> "" Then
          gSL_Select = "Select * from ïiî‘çﬁóøÉrÉÖÅ[ "
          gSL_Select = gSL_Select & " Where ÉAÉhÉåÉX = '" & iB3_ADD & "'"
          
          gSL_Select = gSL_Select & "   AND ïiñ⁄ = 1"    ' 2008/02/06 ADD
          
          If ZTMRead(gSL_Select, 1) Then
             If ZTM.ãÊï™ = 0 Then
                With CisFun
                     .MB_Lines = 6
                     .MB_MSG(2) = "    ìØàÍÇÃÉAÉhÉåÉXÇ™ë∂ç›ÇµÇ‹Ç∑ÅB               "
                     .MB_MSG(4) = "    ïi î‘  ÅF" & ZTM.ï\é¶ïiî‘
                     .MB_MSG(5) = "    éËîzêÊÅF" & ZTM.éÊà¯êÊ
                     .MB_MSG(6) = "    éÛ ì¸  ÅF" & ZTM.éÛì¸
                     .MB_Title = ""
                     .MB_Button = Error
                     .MBOX
                     iB3_ADD.SetFocus
                     Exit Function
                End With
             Else
                If ZTM.çﬁóøä«óùî‘çÜ <> wKanri Or _
                   ZTM.éÊà¯êÊ <> iB2_Torcd Or _
                   ZTM.éÛì¸ <> iB2_Ukeir Then
                   With CisFun
                        .MB_Lines = 8
                        .MB_MSG(2) = "    ìØàÍÇÃÉAÉhÉåÉXÇ™ë∂ç›ÇµÇ‹Ç∑ÅB               "
                        .MB_MSG(4) = "    çﬁ éø ÅF" & ZTM.çﬁéø
                        If ZTM.ê°ñ@ãÊï™ = 0 Then
                           .MB_MSG(5) = "    ê° ñ@ ÅF" & Format(ZTM.î¬å˙, iH2_Itatu.cFormat) & " x " & _
                                                         Format(ZTM.ïù, iH2_Width.cFormat) & " x " & _
                                                         Format(ZTM.í∑Ç≥, iH2_Long.cFormat)
                        End If
                        .MB_MSG(6) = "    éËîzêÊ ÅF" & ZTM.éÊà¯êÊ
                        .MB_MSG(7) = "    éÛ ì¸  ÅF" & ZTM.éÛì¸
                        .MB_Title = ""
                        .MB_Button = Error
                        .MBOX
                   End With
                   iB3_ADD.SetFocus
                   Exit Function
                End If
             End If
          End If
       End If
'    End If                             ' 2009.02.27 del : ïiñ⁄ä÷òA
    
    ' óLå¯ä˙å¿
    If Trim(iB3_SDate) <> "" Then
       If Not iB3_SDate.cChkResult Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    óLå¯ä˙å¿ÅiäJénÅjÇê≥ÇµÇ≠ì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iB3_SDate.SetFocus
          Exit Function
       End If
    End If
    If Trim(iB3_EDate) <> "" Then
       If Not iB3_EDate.cChkResult Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    óLå¯ä˙å¿ÅièIóπÅjÇê≥ÇµÇ≠ì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iB3_EDate.SetFocus
          Exit Function
       End If
    End If
    If Trim(iB3_SDate) <> "" And Trim(iB3_EDate) <> "" Then
       If iB3_SDate > iB3_EDate Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    óLå¯ä˙å¿Çê≥ÇµÇ≠ì¸óÕÇµÇƒâ∫Ç≥Ç¢ÅiäJénÅÑèIóπÅj            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iB3_SDate.SetFocus
          Exit Function
       End If
    End If

    With CisComboTH
        .CodeGet
        wsTHKbn = .Code
    End With
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT STRAT
' ===================================================== *** 2008/06/27 Start
'    If wsTHKbn = "3" And RTrim(iB1_JHinbn) = "" Then
'       With CisFun
'            .MB_Lines = 5
'            .MB_MSG(2) = "     éËîzãÊï™Ç™ÅyéÛããÅzÇÃèÍçá            "
'            .MB_MSG(4) = "     éÛããïiî‘Çì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
'            .MB_Title = ""
'            .MB_Button = Error
'            .MBOX
'       End With
'       B3Comb_THKbn.SetFocus
'       Exit Function
'    End If
    
'ìyäÚí√
'    If wsTHKbn = "3" And RTrim(iB3_JHinbn) = "" Then
'       With CisFun
'            .MB_Lines = 5
'            .MB_MSG(2) = "     éËîzãÊï™Ç™ÅyéÛããÅzÇÃèÍçá            "
'            .MB_MSG(4) = "     éÛããïiî‘(∑∞∫∞ƒﬁ)Çì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
'            .MB_Title = ""
'            .MB_Button = Error
'            .MBOX
'       End With
'       B3Comb_THKbn.SetFocus
'       Exit Function
'    End If
' ===================================================== *** 2008/06/27 End
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
    ' ÉÅÅ[ÉJÅ[
    If Trim(iB3_MTorcd) <> "" Then
        B3lb_MTorcd = ""
        
        If TorNmGet(iB3_MTorcd, 8) Then
            B3lb_MTorcd = TRM_RName
        Else
            With CisFun
                .MB_Lines = 5
                If wsTHKbn = "0" Then
                    .MB_MSG(2) = "    ÉÅÅ[ÉJÅ[ ñ¢ìoò^            "
                Else
                    .MB_MSG(2) = "     î[ì¸êÊ(íºëó) ñ¢ìoò^            "
                End If
                .MB_MSG(4) = "    ÅyéÊà¯êÊÉ}ÉXÉ^Åz           "
                .MB_MSG(2) = "    î[ì¸êÊñ¢ìoò^            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
           End With
           iB3_MTorcd.SetFocus
           Exit Function
        End If
    End If
     
    If Trim(iB3_MTorcd) = "" And _
       Trim(iB3_MUkeir) <> "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    î[ì¸êÊéÛì¸Çì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            If wsTHKbn = "0" Then
                .MB_MSG(2) = "    ÉÅÅ[ÉJÅ[Çì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            Else
                .MB_MSG(2) = "    î[ì¸êÊ(íºëó)Çì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            End If
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB3_MTorcd.SetFocus
        Exit Function
    End If
    ' éÛì¸
    If Trim(iB3_MUkeir) <> "" Then
        gSL_Select = " SELECT * "
        gSL_Select = gSL_Select & " FROM éÊà¯êÊéÛì¸É}ÉXÉ^ "
        gSL_Select = gSL_Select & " WHERE éÊà¯êÊCD = '" & RTrim(iB3_MTorcd) & "'"
        gSL_Select = gSL_Select & "   And éÛì¸ = '" & RTrim(iB3_MUkeir) & "'"
        If Not TUMRead(gSL_Select) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    î[ì¸êÊéÛì¸Ç™ìoò^Ç≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÅB            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iB3_MUkeir.SetFocus
           Exit Function
        End If
    End If
    
'-------------------------------------------------------------------------------' 2009.02.26 add start
    ' â¡çHêÊ
    If Trim(iB3_STorcd) <> "" Then
        B3lb_STorcd = ""
        
        If TorNmGet(iB3_STorcd, 7) Then
            B3lb_STorcd = TRM_RName
        Else
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    â¡çHêÊñ¢ìoò^            "
                .MB_MSG(4) = "    ÅyéÊà¯êÊÉ}ÉXÉ^Åz        "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
           End With
           iB3_STorcd.SetFocus
           Exit Function
        End If
    End If
    ' ìæà”êÊ
    If Trim(iB3_TTorcd) <> "" Then
        B3lb_TTorcd = ""
        
        If TorNmGet(iB3_TTorcd, 2) Then
            B3lb_TTorcd = TRM_RName
        Else
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    ìæà”êÊñ¢ìoò^            "
                .MB_MSG(4) = "    ÅyéÊà¯êÊÉ}ÉXÉ^Åz        "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
           End With
           iB3_TTorcd.SetFocus
           Exit Function
        End If
    End If
'-------------------------------------------------------------------------------' 2009.02.26 add end

'====================================================== *** 2008/06/27 Start
    If RTrim(iB3_HHinbn) <> "" Then
       gSL_Select = ""
       gSL_Select = gSL_Select & " SELECT * FROM çﬁóøéÊà¯êÊÉ}ÉXÉ^"
       gSL_Select = gSL_Select & "  WHERE î≠íçïiî‘ = '" & RTrim(iB3_HHinbn) & "'"
       gSL_Select = gSL_Select & "   AND çﬁóøä«óùî‘çÜ <> " & wKanri
       If ZRMRead(gSL_Select, 1) Then
          With CisFun
             .MB_Lines = 5
             .MB_MSG(2) = "    éwíËÇ≥ÇÍÇΩî≠íçïiî‘ÇÕ     "
             .MB_MSG(4) = "     ä˘Ç…ìoò^çœÇ›Ç≈Ç∑ÅB      "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
          End With
          iB3_HHinbn.SetFocus
          Exit Function
       End If
    End If

    If RTrim(iB3_JHinbn) <> "" Then
       gSL_Select = ""
       gSL_Select = gSL_Select & " SELECT * FROM çﬁóøéÊà¯êÊÉ}ÉXÉ^"
       gSL_Select = gSL_Select & "  WHERE éÛããïiî‘ = '" & RTrim(iB3_JHinbn) & "'"
       gSL_Select = gSL_Select & "   AND çﬁóøä«óùî‘çÜ <> " & wKanri
       If ZRMRead(gSL_Select, 1) Then
          With CisFun
             .MB_Lines = 5
             .MB_MSG(2) = "    éwíËÇ≥ÇÍÇΩéÛããïiî‘ÇÕ     "
             .MB_MSG(4) = "     ä˘Ç…ìoò^çœÇ›Ç≈Ç∑ÅB      "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
          End With
          iB3_JHinbn.SetFocus
          Exit Function
       End If
    End If
'====================================================== *** 2008/06/27 End

'---------------------------------------------------------------' 2009.02.26 add start
    If RTrim(iB3_SHinbn) <> "" Then
       gSL_Select = ""
       gSL_Select = gSL_Select & " SELECT * FROM çﬁóøéÊà¯êÊÉ}ÉXÉ^"
       gSL_Select = gSL_Select & "  WHERE é©é–ïiî‘ = '" & RTrim(iB3_SHinbn) & "'"
       gSL_Select = gSL_Select & "   AND çﬁóøä«óùî‘çÜ <> " & wKanri
       If ZRMRead(gSL_Select, 1) Then
          With CisFun
             .MB_Lines = 5
             .MB_MSG(2) = "    éwíËÇ≥ÇÍÇΩé©é–ïiî‘ÇÕ     "
             .MB_MSG(4) = "     ä˘Ç…ìoò^çœÇ›Ç≈Ç∑ÅB      "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
          End With
          iB3_SHinbn.SetFocus
          Exit Function
       End If
    Else
'ìyäÚí√
'        With CisFun
'           .MB_Lines = 4
'           .MB_MSG(2) = "    é©é–ïiî‘Çì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB     "
'           .MB_Title = ""
'           .MB_Button = Error
'           .MBOX
'        End With
'        iB3_SHinbn.SetFocus
'        Exit Function
    End If
'---------------------------------------------------------------' 2009.02.26 add end
    
    ' î≠íçî‰ó¶
    Dim wHiritsu        As Long
    gLong = CisFun.Val2(iB3_HRitu)
    If gLong <> 0 Then
       If gLong > 100 Then
          With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    î≠íçî‰ó¶ÇÕÇPÇOÇOà»â∫Ç…Çƒê›íËÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
          End With
          iB3_HRitu.SetFocus
          Exit Function
       End If
       wHiritsu = gLong


'---<   2004.10.27  MNT  START   >-----------------------------------------------
'
'       With VsGrid1
'            For mIx = 1 To .Rows - 1
'                If Trim(.TextMatrix(mIx, SCol(1))) = "" Then Exit For
'                   wHiritsu = wHiritsu + CisFun.Val2(Replace(.TextMatrix(mIx, SCol(6)), "%", ""))
'                End If
       With VsGridT
            For mIx = 1 To .Rows - 1
                
                If CisFun.Val2(iB2_Torcd.Tag) <> mIx Then
                    If Trim(.TextMatrix(mIx, SColT(1))) = "" Then Exit For
                    wHiritsu = wHiritsu + CisFun.Val2(Replace(.TextMatrix(mIx, SColT(6)), "%", ""))
                End If
'
'---<   2004.10.27  MNT   END    >-----------------------------------------------


            Next mIx
            If wHiritsu <> 100 Then
               With CisFun
                    .MB_Lines = 5
                    .MB_MSG(2) = "    î≠íçî‰ó¶ÇÃçáåvÇ™ÇPÇOÇOÇ≈ÇÕÇ†ÇËÇ‹ÇπÇÒ            "
                    .MB_MSG(4) = "    Å@Å@Å@Å@Å@ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH            "
                    .MB_Title = ""
                    .MB_Button = Yes_No
                    If Not .MBOX Then
                       iB3_HRitu.SetFocus
                       Exit Function
                    End If
               End With
            End If
       End With
    End If
    
    ' ÉçÉbÉgá@
    If CisFun.Val2(iB3_Lot1) <> 0 Then
       If CisFun.Val2(iB3_Syuyo1) <> 0 Then

'---<   2004.10.27  MNT  START   >-----------------------------------------------
'
'          If CisFun.Val2(iB3_Lot1) Mod CisFun.Val2(iB3_Syuyo1) <> 0 Then
          If (CisFun.Val2(iB3_Lot1) * 1000) Mod _
             (CisFun.Val2(iB3_Syuyo1) * 1000) <> 0 Then
'
'---<   2004.10.27  MNT   END    >-----------------------------------------------

             With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ÉçÉbÉgá@ÇÕé˚óeêîá@ÇÃî{êîÇ≈Ç»ÇØÇÍÇŒÇ»ÇËÇ‹ÇπÇÒÅB            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB3_Lot1.SetFocus
            Exit Function
          End If
       End If
    End If
    ' ÉçÉbÉgáA
    If CisFun.Val2(iB3_Lot2) <> 0 Then
       If CisFun.Val2(iB3_Syuyo2) <> 0 Then

'---<   2004.10.27  MNT  START   >-----------------------------------------------
'
'          If CisFun.Val2(iB3_Lot2) Mod CisFun.Val2(iB3_Syuyo2) <> 0 Then
          If (CisFun.Val2(iB3_Lot2) * 1000) Mod _
             (CisFun.Val2(iB3_Syuyo2) * 1000) <> 0 Then
'
'---<   2004.10.27  MNT   END    >-----------------------------------------------

             With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ÉçÉbÉgáAÇÕé˚óeêîáAÇÃî{êîÇ≈Ç»ÇØÇÍÇŒÇ»ÇËÇ‹ÇπÇÒÅB            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB3_Lot2.SetFocus
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
         iB3_Lot2.SetFocus
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
'-----------------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA start
''-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´-Å´ 2008/02/06 INSERT START
'    If B2Op_Hinmoku(0).Value Then
'        '============== éÛíçï™ÇÃì`ï[éÌóﬁÉ`ÉFÉbÉN ===============
'        SaveKeitai = ""
'        If RTrim(iB3_DensyJ) <> "" Then
'            gSL_Select = "SELECT * FROM ñºèÃÉ}ÉXÉ^ "
'            gSL_Select = gSL_Select & " WHERE ãÊï™ñºèÃ = 'ì`ï[éÌóﬁéÛíç' "
'            gSL_Select = gSL_Select & "   AND íl = '" & RTrim(iB3_DensyJ) & "'"
'            gSL_Select = gSL_Select & "   AND ãÊï™É^ÉCÉv = '' "
'            gSL_Select = gSL_Select & "   AND êîéö1 = 0 "
'            If Not SYMRead(gSL_Select, 1) Then
'               With CisFun
'                    .MB_Lines = 4
'                    .MB_MSG(2) = "    çﬁóøÇ≈ê›íËÇÃçsÇ¶Ç»Ç¢ì`ï[éÌóﬁÇ™ê›íËÇ≥ÇÍÇƒÇ¢Ç‹Ç∑ÅB           "
'                    .MB_Title = ""
'                    .MB_Button = Error
'                    .MBOX
'               End With
'               iB3_Densy.SetFocus
'               Exit Function
'            End If
'            SaveKeitai = CisFun.Left2(SYM.âpêîéö1, 1)
'        Else
'            B3Chk_PrtkbNJ.Value = 1
'            B3Chk_PrtkbSJ.Value = 1
'            B3Chk_PrtkbJJ.Value = 1
'            B3Chk_JRKbnJ.Value = 0
'        End If
'
'        If CisFun.Val2(iB3_PNoJ) > 60 Then
'           With CisFun
'                .MB_Lines = 4
'                .MB_MSG(2) = "    PNO(É|ÉPÉbÉgNO)ÇÕ60à»â∫Ç…Çƒê›íËÇµÇƒâ∫Ç≥Ç¢            "
'                .MB_Title = ""
'                .MB_Button = Error
'                .MBOX
'           End With
'           iB3_PNo.SetFocus
'           Exit Function
'        End If
'    Else
'        '============== î≠íçï™ÇÃì`ï[éÌóﬁÉ`ÉFÉbÉN ===============
''-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™-Å™ 2008/02/06 INSERT END
'-----------------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA end
        SaveKeitai = ""
        
        
        If RTrim(iB3_Densy) <> "" Then
            gSL_Select = "SELECT * FROM ñºèÃÉ}ÉXÉ^ "
            gSL_Select = gSL_Select & " WHERE ãÊï™ñºèÃ = 'ì`ï[éÌóﬁî≠íç' "
            gSL_Select = gSL_Select & "   AND íl = '" & RTrim(iB3_Densy) & "'"
            gSL_Select = gSL_Select & "   AND ãÊï™É^ÉCÉv = '' "
            gSL_Select = gSL_Select & "   AND êîéö1 = 0 "
            If Not SYMRead(gSL_Select, 1) Then
               With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    çﬁóøÇ≈ê›íËÇÃçsÇ¶Ç»Ç¢ì`ï[éÌóﬁÇ™ê›íËÇ≥ÇÍÇƒÇ¢Ç‹Ç∑ÅB           "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
               End With
               iB3_Densy.SetFocus
               Exit Function
            End If
            SaveKeitai = CisFun.Left2(SYM.âpêîéö1, 1)
            If (RTrim(iB3_Densy) = "3" Or RTrim(iB3_Densy) = "4") And (B3Chk_PrtkbS.Value = 0 Or B3Chk_PrtkbJ.Value = 0) Then
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
            B3Chk_PrtkbN.Value = 1
            B3Chk_PrtkbS.Value = 1
            B3Chk_PrtkbJ.Value = 1
            B3Chk_JRKbn.Value = 0
        End If
        
        If CisFun.Val2(iB3_PNo) > 60 Then
           With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    PNO(É|ÉPÉbÉgNO)ÇÕ60à»â∫Ç…Çƒê›íËÇµÇƒâ∫Ç≥Ç¢            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
           End With
           iB3_PNo.SetFocus
           Exit Function
        End If
'    End If                         ' 2009.02.27 del : ïiñ⁄ä÷òA
    
    
    '===================================================== 2010/01/16 Start
    If B3Chk_Tana.Value = 0 Then
       gSL_Select = ""
       gSL_Select = gSL_Select & " SELECT * FROM çﬁóøéÊà¯êÊÉ}ÉXÉ^"
       gSL_Select = gSL_Select & "  WHERE çﬁóøä«óùî‘çÜ = " & wKanri
       gSL_Select = gSL_Select & "    AND RTRIM(éÊà¯êÊ) + RTRIM(éÛì¸) <> '" & RTrim(iB2_Torcd) & RTrim(iB2_Ukeir) & "'"
       gSL_Select = gSL_Select & "    AND ISNULL(íIâµãÊï™,0) = 0 "
       If ZTMRead(gSL_Select, 1) Then
           With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ä˘Ç…ëºéÊà¯êÊÇ≈[íIâµãÊï™ = Ç∑ÇÈ]Ç™ê›íËÇ≥ÇÍÇƒÇ¢Ç‹Ç∑    "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
           End With
           B3Chk_Tana.SetFocus
           Exit Function
       End If
    End If
    '===================================================== 2010/01/16 End
    
    
    Body3Chk = True
End Function
'***********************************************************************************
'*
'*       ÉfÅ[É^ÉxÅ[ÉXì¸óÕ Åï ï\é¶
'*
'***********************************************************************************
'+---------------------------------+
'+      äÓñ{èÓïÒ
'+---------------------------------+
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    Call HeadBodyClear("B")
    
    If Not ZRMRead(gSL_Select, 1) Then
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
        
        iH1_Zaist = RTrim(ZRM.çﬁéø)
        Call Head2Set
        DBInput = True
    End If
DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'***********************************************************************************
'*      HEAD2 ÉfÅ[É^ÉxÅ[ÉXì¸óÕ Åï ï\é¶
'***********************************************************************************
Private Function Head2Input(InputKey As String)

    Head2Input = False
    On Error GoTo Head2Input_Err:
    
    If InputKey = "INV" Then
        gSL_Select = " SELECT * "
        gSL_Select = gSL_Select & " FROM çﬁóøÉ}ÉXÉ^ "
        gSL_Select = gSL_Select & " WHERE çﬁéø = '" & RTrim(iH1_Zaist) & "'"
    End If
    
    If Not ZRMRead(gSL_Select, 1) Then
        If InputKey = "Back" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     êÊì™ÉåÉRÅ[ÉhÇ≈Ç∑            "
                .MB_Title = "Åyå„ëﬁÅz"
                .MB_Button = OK
                .MBOX
            End With
            GoTo Head2Input_Ed
        End If
        If InputKey = "Next" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     ç≈èIÉåÉRÅ[ÉhÇ≈Ç∑            "
                .MB_Title = "ÅyëOêiÅz"
                .MB_Button = OK
                .MBOX
            End With
            GoTo Head2Input_Ed
        End If
    Else
        Call HeadBodyClear("H2")
        Call HeadBodyClear("B")
        
        iH1_Zaist = RTrim(ZRM.çﬁéø)
        H1Chk_Size.Value = ZRM.ê°ñ@ãÊï™
        
        Call Head2Set
        Head2Input = True
    End If
Head2Input_Ed:
    Exit Function
Head2Input_Err:
     CisFun.ErrorBox
End Function
'**************************************************************************************
'*      HEAD2 ì‡óeæØƒèàóù
'**************************************************************************************
Private Sub Head2Set()
    
    Call CisComboZK.ComboSet
    gSL_Select = "SELECT * FROM çﬁóøÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " WHERE çﬁéø = '" & RTrim(iH1_Zaist) & "'"
    gSL_Select = gSL_Select & " ORDER BY î¬å˙,ïù,í∑Ç≥"
    If H1Chk_Size.Value = 0 Then
        If ZRMRead(gSL_Select) Then
            GridSet
        End If
    Else
        If ZRMRead(gSL_Select, 1) Then
            iB1_Name = RTrim(ZRM.çﬁóøñº)
            iB1_Zaisy = RTrim(ZRM.çﬁéÌ)
            Call iB1_Zaisy_LostFocus
            iB1_Keijo = RTrim(ZRM.å`èÛ)
            Call iB1_Keijo_LostFocus
            iB1_Tani = RTrim(ZRM.íPà )
            Call iB1_Tani_LostFocus
            iB1_Kansn = CisFun.RSetFld(ZRM.ä∑éZíl, 10, "###0.000")
            iB1_TTani = RTrim(ZRM.éËîzíPà )
            Call iB1_TTani_LostFocus
        
            Call CisComboZK.ComboSet
            With CisComboZK
                .Code = ZRM.ç›å…ä«óùãÊï™
                .TextGet
            End With
            iB1_Zaiko = CisFun.RSetFld(ZRM.ç›å…ä«óùåWêî, 0, iB1_Zaiko.cFormat)
            iB1_ZaikoDay = CisFun.RSetFld(ZRM.ç›å…ä«óùì˙êî, 0, iB1_ZaikoDay.cFormat)
            iB1_ZaikoL = CisFun.RSetFld(ZRM.ç›å…ä«óùî≠íçì_, 0, iB1_ZaikoL.cFormat)
            B1Chk_JSKbn.Value = ZRM.é¿êîãÊï™
            B1Chk_ZKKbn.Value = ZRM.ç›å…ãÊï™
            
'====================================================== *** 2008/06/27 Start(Comment)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
'            iB1_JHinbn = RTrim(ZRM.éÛããïiî‘)
'            iB1_HHinbn = RTrim(ZRM.î≠íçïiî‘)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'====================================================== *** 2008/06/27 End(Comment)
            B1Op_Syoyo(CisFun.Val2(ZRM.èäóvó ãÊï™)).Value = True                ' 2009.02.26 add
            iB1_Code = ZRM.çﬁéÌÉRÅ[Éh                                           ' 2009.02.27 add
            
            wKanri = ZRM.çﬁóøä«óùî‘çÜ
        End If
    End If
End Sub
'********************************************************************
'*                                                                  *
'*          ÉfÅ[É^ÉxÅ[ÉXçXêV                                         *
'*                                                                  *
'********************************************************************
'+-------------------------------------------------------+
'+      çÌ èú èà óù (çﬁóø)
'+-------------------------------------------------------+
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
        If H1Chk_Size.Value = 0 Then
            wKanri = CisFun.Val2(vsGrid1.TextMatrix(vsGrid1.Row, 4))
        Else
            gSL_Select = "SELECT * FROM çﬁóøÉ}ÉXÉ^"
            gSL_Select = gSL_Select & " WHERE çﬁéø = '" & RTrim(iH1_Zaist) & "'"
            Call ZRMRead(gSL_Select, 1)
            wKanri = ZRM.çﬁóøä«óùî‘çÜ
        End If
            
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/03/02 INSERT START
'    ' òAågçﬁóøÉ}ÉXÉ^(çÌèú)ìoò^
'        .SQL = ""
'        .SQL = .SQL & " INSERT INTO òAågçﬁóøÉ}ÉXÉ^"
'        .SQL = .SQL & " SELECT é©é–ïiî‘, '', '', ZRM.çﬁóøñº, '', ZRM.ï\é¶ê°ñ@, '', '',"
'        .SQL = .SQL & "        ZRM.ç›å…ä«óùãÊï™, ZRM.ç›å…ä«óùêîíl, ZRM.ç›å…ä«óùåWêî, ZRM.ç›å…ä«óùì˙êî, ZRM.ç›å…ä«óùî≠íçì_,"
'        .SQL = .SQL & "        '','','','','','',0,'','',"
'        .SQL = .SQL & "        ZRM.âºãÊï™,ZRM.çÏê¨ì˙,ZRM.çÏê¨é“,ZRM.çÏê¨í[ññ,ZRM.çXêVì˙,ZRM.çXêVé“,ZRM.çXêVí[ññ,"
'        .SQL = .SQL & "        ZRM.ê›íËêî,ZRM.çﬁéø,ZRM.ê°ñ@ãÊï™, ZRM.î¬å˙, ZRM.ïù, ZRM.í∑Ç≥,"
'        .SQL = .SQL & "        ZRM.çﬁéÌ, ZRM.å`èÛ, ZRM.íPà , ZRM.ä∑éZíl, ZRM.éËîzíPà , ZRM.çﬁéÌÉRÅ[Éh, ZRM.çﬁóøä«óùî‘çÜ, ZRM.èäóvó ãÊï™,"
'        .SQL = .SQL & "        CONVERT(VARCHAR(8),GETDATE(),112) + REPLACE(CONVERT(VARCHAR(8), GETDATE(),108),':',''), 1, 1"
'        .SQL = .SQL & "   FROM çﬁóøéÊà¯êÊÉ}ÉXÉ^ ZTM"
'        .SQL = .SQL & "        LEFT OUTER JOIN çﬁóøÉ}ÉXÉ^ ZRM"
'        .SQL = .SQL & "             ON ZTM.çﬁóøä«óùî‘çÜ = ZRM.çﬁóøä«óùî‘çÜ"
'        .SQL = .SQL & "  WHERE ZTM.çﬁóøä«óùî‘çÜ = " & ZRM.çﬁóøä«óùî‘çÜ & ""
'
'        On Error GoTo DBDelete_Err
'        .DBExec
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/03/02 INSERT END
            
    ' çﬁóøÉ}ÉXÉ^
        .SQL = "DELETE FROM çﬁóøÉ}ÉXÉ^ "
        .SQL = .SQL & " WHERE çﬁóøä«óùî‘çÜ = " & wKanri
        .DBExec

    ' çﬁóøéÊà¯êÊÉ}ÉXÉ^
        .SQL = "DELETE FROM çﬁóøéÊà¯êÊÉ}ÉXÉ^ "
        .SQL = .SQL & " WHERE çﬁóøä«óùî‘çÜ = " & wKanri
        On Error GoTo DBDelete_Err
        .DBExec
    
    ' ïiî‘éÊà¯çÏã∆É}ÉXÉ^
        .SQL = "DELETE FROM çﬁóøïiî‘É}ÉXÉ^ "
        .SQL = .SQL & " WHERE çﬁóøä«óùî‘çÜ = " & wKanri
        On Error GoTo DBDelete_Err
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
'+-------------------------------------------------------+
'+      çÌ èú èà óù (éÊà¯êÊ)
'+-------------------------------------------------------+
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
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/03/02 INSERT START
'    ' òAågçﬁóøÉ}ÉXÉ^(çÌèú)ìoò^
'        .SQL = ""
'        .SQL = .SQL & " INSERT INTO òAågçﬁóøÉ}ÉXÉ^"
'        .SQL = .SQL & " SELECT é©é–ïiî‘, '', '', ZRM.çﬁóøñº, '', ZRM.ï\é¶ê°ñ@, '', '',"
'        .SQL = .SQL & "        ZRM.ç›å…ä«óùãÊï™, ZRM.ç›å…ä«óùêîíl, ZRM.ç›å…ä«óùåWêî, ZRM.ç›å…ä«óùì˙êî, ZRM.ç›å…ä«óùî≠íçì_,"
'        .SQL = .SQL & "        '','','','','','',0,'','',"
'        .SQL = .SQL & "        ZRM.âºãÊï™,ZRM.çÏê¨ì˙,ZRM.çÏê¨é“,ZRM.çÏê¨í[ññ,ZRM.çXêVì˙,ZRM.çXêVé“,ZRM.çXêVí[ññ,"
'        .SQL = .SQL & "        ZRM.ê›íËêî,ZRM.çﬁéø,ZRM.ê°ñ@ãÊï™, ZRM.î¬å˙, ZRM.ïù, ZRM.í∑Ç≥,"
'        .SQL = .SQL & "        ZRM.çﬁéÌ, ZRM.å`èÛ, ZRM.íPà , ZRM.ä∑éZíl, ZRM.éËîzíPà , ZRM.çﬁéÌÉRÅ[Éh, ZRM.çﬁóøä«óùî‘çÜ, ZRM.èäóvó ãÊï™,"
'        .SQL = .SQL & "        CONVERT(VARCHAR(8),GETDATE(),112) + REPLACE(CONVERT(VARCHAR(8), GETDATE(),108),':',''), 1, 1"
'        .SQL = .SQL & "   FROM çﬁóøéÊà¯êÊÉ}ÉXÉ^ ZTM"
'        .SQL = .SQL & "        LEFT OUTER JOIN çﬁóøÉ}ÉXÉ^ ZRM"
'        .SQL = .SQL & "             ON ZTM.çﬁóøä«óùî‘çÜ = ZRM.çﬁóøä«óùî‘çÜ"
'        .SQL = .SQL & "  WHERE ZTM.çﬁóøä«óùî‘çÜ = " & wKanri & ""
'        .SQL = .SQL & "    AND ZTM.éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
'        .SQL = .SQL & "    AND ZTM.éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
'
'        On Error GoTo DBDeleteTor_Err
'        .DBExec
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/03/02 INSERT END
        .SQL = "DELETE FROM çﬁóøéÊà¯êÊÉ}ÉXÉ^ "
        .SQL = .SQL & " Where çﬁóøä«óùî‘çÜ = " & wKanri
        .SQL = .SQL & " And   éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
        .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
        B2Op_Hinmoku(1).Value = True                                            ' 2009.02.27 add : ïiñ⁄ä÷òA
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
        If B2Op_Hinmoku(0).Value Then
           .SQL = .SQL & " And   ïiñ⁄ = 0"
        Else
           .SQL = .SQL & " And   ïiñ⁄ = 1"
        End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
        On Error GoTo DBDeleteTor_Err
        .DBExec
        
        .DBTran (TransCommit) ' ƒ◊›ªﬁ∏ºÆ› èIóπ
        
        gLong = CisFun.Val2(iB2_Torcd.Tag)
        With VsGridT
             .RemoveItem gLong
             If .Rows < CisVsGrid.DispMax Then
                .Rows = CisVsGrid.DispMax
             End If
        End With
        CisVsGridT.ReNumber
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
Private Function DBPut(KB As Boolean)
    DBPut = False
    Dim AddFlg  As Boolean

    Static ws_Size1     As String
    Static ws_Size2     As String
    Static ws_Size3     As String

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    If KB Then
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
                iB1_Name.SetFocus
                Exit Function
            End If
        End With
    End If
    On Error GoTo DBPut_Err

'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/04/01 UPDATE START
    'If SyoriKB = "ADD" Then
    If SyoriKB = "ADD" And wKanri = 0 Then
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/04/01 UPDATE END
        With CisDB
            .SQL = "î‘çÜälìæ"
            .StoadoCount = 3
            .ParaNo = 0: .ParaIO = Returnóp
            .ParaNo = 1: .ParaIO = Inputóp
            .ParaNo = 2: .ParaIO = OutPutóp
            .ParaNo = 3: .ParaIO = Returnóp
        
            .ParaNo = 1: .ParaValue = 1
        
            If Not .DBStored Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "     çﬁóøä«óùî‘çÜ çÃî‘Ç…é∏îsÇµÇ‹ÇµÇΩ    "
                    .MB_Title = "çÃî‘èàóù"
                    .MB_Button = Error
                    If Not .MBOX Then GoTo DBPut_Err
                End With
            Else
                .ParaNo = 2: wKanri = .ParaValue
            End If
        End With
    End If

'   ÉTÉCÉXïœçXÉ`ÉFÉbÉN
    gSL_Select = "SELECT * FROM ñºèÃÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " WHERE ãÊï™ñºèÃ = 'å`èÛãÊï™' "
    gSL_Select = gSL_Select & "   AND íl = '" & RTrim(iB1_Keijo) & "'"
    gSL_Select = gSL_Select & "   AND ãÊï™É^ÉCÉv = '' "
    Call SYMRead(gSL_Select, 1)
    ws_Size1 = SYM.âpêîéö1
    ws_Size2 = SYM.âpêîéö2
    ws_Size3 = SYM.âpêîéö3

'   É}ÉXÉ^çÄñ⁄à⁄ëó
    gSL_Select = "SELECT * From çﬁóøÉ}ÉXÉ^"
    gSL_Select = gSL_Select & " WHERE çﬁéø = '" & RTrim(iH1_Zaist) & "'"
    gSL_Select = gSL_Select & "   AND î¬å˙ =  " & CisFun.Val2(iH2_Itatu)
    gSL_Select = gSL_Select & "   AND ïù   =  " & CisFun.Val2(iH2_Width)
    gSL_Select = gSL_Select & "   AND í∑Ç≥   =  " & CisFun.Val2(iH2_Long)
    If ZRMRead(gSL_Select, 1) Then
        AddFlg = False
    Else
        Call ItemsClearZRM
        AddFlg = True
        ZRM.çﬁéø = RTrim(iH1_Zaist)

        ZRM.î¬å˙ = CisFun.Val2(iH2_Itatu)
        ZRM.ïù = CisFun.Val2(iH2_Width)
        ZRM.í∑Ç≥ = CisFun.Val2(iH2_Long)
        ZRM.ê°ñ@ãÊï™ = CisFun.Val2(H1Chk_Size.Value)
    End If
    With ZRM
        .çﬁóøä«óùî‘çÜ = wKanri
        .çﬁóøñº = RTrim(iB1_Name)
        .çﬁéÌ = RTrim(iB1_Zaisy)
        .å`èÛ = RTrim(iB1_Keijo)
        .çﬁéÌÉRÅ[Éh = RTrim(iB1_Code)
        .íPà  = RTrim(iB1_Tani)
        .ä∑éZíl = CisFun.Val2(iB1_Kansn)
        .éËîzíPà  = RTrim(iB1_TTani)
        .ç›å…ä«óùãÊï™ = CisFun.Val2(B1Chk_ZKKbn)
        .é¿êîãÊï™ = B1Chk_JSKbn.Value
        .ç›å…ãÊï™ = B1Chk_ZKKbn.Value

        gStr = ""
        If .ê°ñ@ãÊï™ = 0 Then
            If RTrim(ws_Size1) = "" Then
                If .î¬å˙ = 0 Then
                    gStr = Space(7)
                Else
                    gStr = Space(7 - Len(Format(.î¬å˙, "###0.00"))) & Format(.î¬å˙, "###0.00")
                End If
            Else
                If .î¬å˙ = 0 Then
                    gStr = RTrim(Mid(ws_Size1, 1, 7)) & Space(7 - Len(RTrim(Mid(ws_Size1, 1, 7))))
                Else
                    gStr = Space(7 - Len(Format(.î¬å˙, "###0.00"))) & Format(.î¬å˙, "###0.00")
                End If
            End If
            If RTrim(ws_Size2) = "" Then
                If .ïù = 0 Then
                    gStr = gStr & "x" & Space(7)
                Else
                    gStr = gStr & "x" & Space(7 - Len(Format(.ïù, "###0.00"))) & Format(.ïù, "###0.00")
                End If
            Else
                If .ïù = 0 Then
                    gStr = gStr & "x" & RTrim(Mid(ws_Size2, 1, 7)) & Space(7 - Len(RTrim(Mid(ws_Size2, 1, 7))))
                Else
                    gStr = gStr & "x" & Space(7 - Len(Format(.ïù, "###0.00"))) & Format(.ïù, "###0.00")
                End If
            End If
            If RTrim(ws_Size3) = "" Then
                If .í∑Ç≥ = 0 Then
                    gStr = gStr & "x" & Space(7)
                Else
                    gStr = gStr & "x" & Space(7 - Len(Format(.í∑Ç≥, "###0.00"))) & Format(.í∑Ç≥, "###0.00")
                End If
            Else
                If .í∑Ç≥ = 0 Then
                    gStr = gStr & "x" & RTrim(Mid(ws_Size3, 1, 7)) & Space(7 - Len(RTrim(Mid(ws_Size3, 1, 7))))
                Else
                    gStr = gStr & "x" & Space(7 - Len(Format(.í∑Ç≥, "###0.00"))) & Format(.í∑Ç≥, "###0.00")
                End If
            End If
        End If
        .ï\é¶ê°ñ@ = gStr
'
        With CisComboZK
            .CodeGet
            ZRM.ç›å…ä«óùãÊï™ = CisFun.Val2(.Code)
        End With
        .ç›å…ä«óùåWêî = CisFun.Val2(iB1_Zaiko)
        .ç›å…ä«óùî≠íçì_ = CisFun.Val2(iB1_ZaikoL)
        .ç›å…ä«óùì˙êî = CisFun.Val2(iB1_ZaikoDay)
        Select Case .ç›å…ä«óùãÊï™
             Case 2: .ç›å…ä«óùêîíl = .ç›å…ä«óùåWêî
             Case 3: .ç›å…ä«óùêîíl = .ç›å…ä«óùî≠íçì_
             Case 4: .ç›å…ä«óùêîíl = .ç›å…ä«óùì˙êî
        End Select
        
        .âºãÊï™ = 0
        
'====================================================== *** 2008/06/27 Start(Comment)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
'        .éÛããïiî‘ = RTrim(iB1_JHinbn)
'        .î≠íçïiî‘ = RTrim(iB1_HHinbn)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'====================================================== *** 2008/06/27 End(Comment)

        .èäóvó ãÊï™ = B1Op_Syoyo(0).Tag

    End With
    Call CisDB.DBTran(TransBegin)  ' ƒ◊›ªﬁ∏ºÆ›äJén
' ( í«â¡èàóù )
    If AddFlg Then
        ZRM.çÏê¨é“ = gTanto
        Call ZRMInsert
        iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
    Else
' ( çXêVèàóù )
        ZRM.çXêVé“ = gTanto
            
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/03/02 INSERT START
'    ' òAågçﬁóøÉ}ÉXÉ^(çÌèú)ìoò^
'        With CisDB
'            .SQL = ""
'            .SQL = .SQL & " INSERT INTO òAågçﬁóøÉ}ÉXÉ^"
'            .SQL = .SQL & " SELECT é©é–ïiî‘, '', '', ZRM.çﬁóøñº, '', ZRM.ï\é¶ê°ñ@, '', '',"
'            .SQL = .SQL & "        ZRM.ç›å…ä«óùãÊï™, ZRM.ç›å…ä«óùêîíl, ZRM.ç›å…ä«óùåWêî, ZRM.ç›å…ä«óùì˙êî, ZRM.ç›å…ä«óùî≠íçì_,"
'            .SQL = .SQL & "        '','','','','','',0,'','',"
'            .SQL = .SQL & "        ZRM.âºãÊï™,ZRM.çÏê¨ì˙,ZRM.çÏê¨é“,ZRM.çÏê¨í[ññ,ZRM.çXêVì˙,ZRM.çXêVé“,ZRM.çXêVí[ññ,"
'            .SQL = .SQL & "        ZRM.ê›íËêî,ZRM.çﬁéø,ZRM.ê°ñ@ãÊï™, ZRM.î¬å˙, ZRM.ïù, ZRM.í∑Ç≥,"
'            .SQL = .SQL & "        ZRM.çﬁéÌ, ZRM.å`èÛ, ZRM.íPà , ZRM.ä∑éZíl, ZRM.éËîzíPà , ZRM.çﬁéÌÉRÅ[Éh, ZRM.çﬁóøä«óùî‘çÜ, ZRM.èäóvó ãÊï™,"
'            .SQL = .SQL & "        CONVERT(VARCHAR(8),GETDATE(),112) + REPLACE(CONVERT(VARCHAR(8), GETDATE(),108),':',''), 1, 1"
'            .SQL = .SQL & "   FROM çﬁóøéÊà¯êÊÉ}ÉXÉ^ ZTM"
'            .SQL = .SQL & "        LEFT OUTER JOIN çﬁóøÉ}ÉXÉ^ ZRM"
'            .SQL = .SQL & "             ON ZTM.çﬁóøä«óùî‘çÜ = ZRM.çﬁóøä«óùî‘çÜ"
'            .SQL = .SQL & "  WHERE ZTM.çﬁóøä«óùî‘çÜ = " & ZRM.çﬁóøä«óùî‘çÜ & ""
'
'            .DBExec
'        End With
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/03/02 INSERT END
        Call ZRMUpdate
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/03/02 INSERT START
'    ' òAågçﬁóøÉ}ÉXÉ^(çÌèú)ìoò^
'        With CisDB
'            .SQL = ""
'            .SQL = .SQL & " INSERT INTO òAågçﬁóøÉ}ÉXÉ^"
'            .SQL = .SQL & " SELECT é©é–ïiî‘, '', '', ZRM.çﬁóøñº, '', ZRM.ï\é¶ê°ñ@, '', '',"
'            .SQL = .SQL & "        ZRM.ç›å…ä«óùãÊï™, ZRM.ç›å…ä«óùêîíl, ZRM.ç›å…ä«óùåWêî, ZRM.ç›å…ä«óùì˙êî, ZRM.ç›å…ä«óùî≠íçì_,"
'            .SQL = .SQL & "        '','','','','','',0,'','',"
'            .SQL = .SQL & "        ZRM.âºãÊï™,ZRM.çÏê¨ì˙,ZRM.çÏê¨é“,ZRM.çÏê¨í[ññ,ZRM.çXêVì˙,ZRM.çXêVé“,ZRM.çXêVí[ññ,"
'            .SQL = .SQL & "        ZRM.ê›íËêî,ZRM.çﬁéø,ZRM.ê°ñ@ãÊï™, ZRM.î¬å˙, ZRM.ïù, ZRM.í∑Ç≥,"
'            .SQL = .SQL & "        ZRM.çﬁéÌ, ZRM.å`èÛ, ZRM.íPà , ZRM.ä∑éZíl, ZRM.éËîzíPà , ZRM.çﬁéÌÉRÅ[Éh, ZRM.çﬁóøä«óùî‘çÜ, ZRM.èäóvó ãÊï™,"
'            .SQL = .SQL & "        CONVERT(VARCHAR(8),GETDATE(),112) + REPLACE(CONVERT(VARCHAR(8), GETDATE(),108),':',''), 2, 1"
'            .SQL = .SQL & "   FROM çﬁóøéÊà¯êÊÉ}ÉXÉ^ ZTM"
'            .SQL = .SQL & "        LEFT OUTER JOIN çﬁóøÉ}ÉXÉ^ ZRM"
'            .SQL = .SQL & "             ON ZTM.çﬁóøä«óùî‘çÜ = ZRM.çﬁóøä«óùî‘çÜ"
'            .SQL = .SQL & "  WHERE ZTM.çﬁóøä«óùî‘çÜ = " & ZRM.çﬁóøä«óùî‘çÜ & ""
'
'            .DBExec
'        End With
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/03/02 INSERT END
        iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
    End If

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
    
' äÓñ{èÓïÒçÏê¨
    Call DBPut(False)
    
    With VsGridT
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
        
        ' ±ƒﬁ⁄Ωè»ó™éû,±ƒﬁ⁄Ωé©ìÆçÃî‘
        wAdress = ""
        If CisFun.Val2(iB3_ADD) = 0 And B3Chk_Auto.Value = 1 Then
           gSL_Select = "Select Max(ÉAÉhÉåÉX) ÉAÉhÉåÉX From ïiî‘çﬁóøÉrÉÖÅ["
           Call ZTMRead(gSL_Select, 1)
           If Trim(ZTM.ÉAÉhÉåÉX) = "" Then
              wAdress = Format(1, iB3_ADD.cFormat)
           Else
              gStr = String(iB3_ADD.MaxLength, "9")
              gLong = CisFun.Val2(ZTM.ÉAÉhÉåÉX) + 1
              gInt = 0
              Do Until gInt > 1
                 gSL_Select = "Select ÉAÉhÉåÉX From ïiî‘çﬁóøÉrÉÖÅ[ "
                 gSL_Select = gSL_Select & " Where ÉAÉhÉåÉX = '" & Format(gLong, iB3_ADD.cFormat) & "'"
                 If Not ZTMRead(gSL_Select, 1) Then
                    wAdress = Format(gLong, iB3_ADD.cFormat)
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
        
        gSL_Select = "Select * from çﬁóøéÊà¯êÊÉ}ÉXÉ^ "
        gSL_Select = gSL_Select & " Where çﬁóøä«óùî‘çÜ = " & wKanri
        gSL_Select = gSL_Select & " And   éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
        gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
        If ZTMRead(gSL_Select, 1) Then
            AddFlg = False
        Else
            Call ItemsClearZTM
            AddFlg = True
            ZTM.çﬁóøä«óùî‘çÜ = wKanri
            ZTM.éÊà¯êÊ = RTrim(iB2_Torcd)
            ZTM.éÛì¸ = RTrim(iB2_Ukeir)
            ZTM.çﬁéø = RTrim(iH1_Zaist)
            ZTM.ê°ñ@ãÊï™ = H1Chk_Size.Value
            ZTM.î¬å˙ = CisFun.Val2(iH2_Itatu)
            ZTM.ïù = CisFun.Val2(iH2_Width)
            ZTM.í∑Ç≥ = CisFun.Val2(iH2_Long)
        End If
        With ZTM
            
            .îwî‘çÜ = iB3_Seban
            If B3Op_SKKbn(0).Value Then
               .SKãÊï™ = 0
            Else
               .SKãÊï™ = 1
            End If
            If CisFun.Val2(iB3_Cycle1) <> 0 Or CisFun.Val2(iB3_Cycle2) <> 0 Or _
               CisFun.Val2(iB3_Cycle3) <> 0 Then
               .ÉTÉCÉNÉã = Format(iB3_Cycle1, "0") & _
                             Format(iB3_Cycle2, "00") & _
                             Format(iB3_Cycle3, "00")
            End If
            .é˚óeêî1 = CisFun.Val2(iB3_Syuyo1)
            .é˚óeäÌ1 = iB3_Hako1
            .ÉçÉbÉg1 = CisFun.Val2(iB3_Lot1)
            .é˚óeêî2 = CisFun.Val2(iB3_Syuyo2)
            .é˚óeäÌ2 = iB3_Hako2
            .ÉçÉbÉg2 = CisFun.Val2(iB3_Lot2)
            .ïœçXì˙ = iB3_CDate
            If iB3_CBin <> "" Then
               .ïœçXï÷ = Format(CisFun.Val2(iB3_CBin), "00")
            Else
               .ïœçXï÷ = ""
            End If
'========================================== 2009/12/08 Start
'            .íuèÍ = iB3_Okiba
            .íuèÍ = RTrim(iB3_Okiba1) & Space(2 - Len(iB3_Okiba1)) & _
                    RTrim(iB3_Okiba2) & Space(2 - Len(iB3_Okiba2)) & _
                    RTrim(iB3_Okiba3) & Space(2 - Len(iB3_Okiba3)) & _
                    RTrim(iB3_Okiba4) & Space(2 - Len(iB3_Okiba4))
'========================================== 2009/12/08 End
'---------------------------------------------------------------------------------------' 2008.02.27 del ïiñ⁄ä÷òA start
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
'            If B2Op_Hinmoku(0).Value Then
'               .ïiñ⁄ = 0
'            Else
'               .ïiñ⁄ = 1
'            End If
'
'            If B2Op_Hinmoku(0).Value Then
'                .îıçl = RTrim(iB3_BikoJ)
'                .ì`ï[éÌóﬁ = iB3_DensyJ
'                .î[ïièëå`ë‘ = SaveKeitai
'                .éÛóÃèëå`ë‘ = SaveKeitai
'                .î[ïièëî≠çsãÊï™ = B3Chk_PrtkbNJ.Value
'                .éxããèëî≠çsãÊï™ = B3Chk_PrtkbSJ.Value
'                .éÛóÃèëî≠çsãÊï™ = B3Chk_PrtkbJJ.Value
'
'                .éÛóÃèëî≠çs = B3Chk_JRKbn.Value
'
'                If CisFun.Val2(iB3_ADD) = 0 And Trim(wAdress) <> "" Then
'                   ZTM.ÉAÉhÉåÉX = wAdress
'                   iB3_ADD = wAdress
'                Else
'                   ZTM.ÉAÉhÉåÉX = RTrim(iB3_ADD)
'                End If
'            Else
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'---------------------------------------------------------------------------------------' 2008.02.27 del ïiñ⁄ä÷òA end
                With CisComboTH
                    .CodeGet
                    ZTM.éËîzãÊï™ = .Code
                End With
                .ÉÅÅ[ÉJÅ[ = iB3_MTorcd
                .ÉÅÅ[ÉJÅ[éÛì¸ = iB3_MUkeir
                .å_ñÒêÊ = iB3_KTorcd
                .åWêî1 = CisFun.Val2(iB3_Keisu1)
                .åWêî2 = CisFun.Val2(iB3_Keisu2)
                .åWêî3 = CisFun.Val2(iB3_Keisu3)
                .LT = CisFun.Val2(iB3_LT)
                .Week1 = B3Chk_Week(0).Value
                .Week2 = B3Chk_Week(1).Value
                .Week3 = B3Chk_Week(2).Value
                .Week4 = B3Chk_Week(3).Value
                .Week5 = B3Chk_Week(4).Value
                .Day1 = CisFun.Val2(iB3_Day(0))
                .Day2 = CisFun.Val2(iB3_Day(1))
                .Day3 = CisFun.Val2(iB3_Day(2))
                .Day4 = CisFun.Val2(iB3_Day(3))
                .Day5 = CisFun.Val2(iB3_Day(4))
                .î≠íçî‰ó¶ = CisFun.Val2(iB3_HRitu)
                .óLå¯ä˙å¿äJén = iB3_SDate
                .óLå¯ä˙å¿èIóπ = iB3_EDate
                .ì‡é¶î≠íçãÊï™ = B3Chk_NHKbn.Value
                .îıçl = iB3_Biko
                
                If CisFun.Val2(iB3_ADD) = 0 And Trim(wAdress) <> "" Then
                   ZTM.ÉAÉhÉåÉX = wAdress
                   iB3_ADD = wAdress
                Else
                   ZTM.ÉAÉhÉåÉX = RTrim(iB3_ADD)
                End If
                .ì`ï[éÌóﬁ = iB3_Densy
                .î[ïièëå`ë‘ = SaveKeitai
                .éÛóÃèëå`ë‘ = SaveKeitai
'''2007/12/11DEL            With CisComboNH
'''2007/12/11DEL                .CodeGet
'''2007/12/11DEL                ZTM.î[ïièëå`ë‘ = .Code
'''2007/12/11DEL            End With
'''2007/12/11DEL            With CisComboJR
'''2007/12/11DEL                .CodeGet
'''2007/12/11DEL                ZTM.éÛóÃèëå`ë‘ = .Code
'''2007/12/11DEL            End With
                
                '=============================================== ***** 2008/06/11 Start
                '''ZTM.î[ïièëî≠çsãÊï™ = B3Chk_PrtkbN.Value
                If ZTM.éËîzãÊï™ <> 4 Then
                   ZTM.î[ïièëî≠çsãÊï™ = B3Chk_PrtkbN.Value
                Else
                   ZTM.î[ïièëî≠çsãÊï™ = B3Chk_PrtkbSN.Value
                End If
                '=============================================== ***** 2008/06/11 End
'---------------------------------------------------------------------------------------' 2009.02.27 del : íºëóì`ï[î≠çsãÊï™ä÷òA start
'                '=============================================== 2009/01/15 Start
'                If ZTM.éËîzãÊï™ = 1 Then
'                   ZTM.íºëóì`ï[î≠çsãÊï™ = B3Chk_PrtkbSN.Value
'                End If
'                '=============================================== 2009/01/15 End
'---------------------------------------------------------------------------------------' 2009.02.27 del : íºëóì`ï[î≠çsãÊï™ä÷òA del
                
                ZTM.éxããèëî≠çsãÊï™ = B3Chk_PrtkbS.Value
                ZTM.éÛóÃèëî≠çsãÊï™ = B3Chk_PrtkbJ.Value
                
                ZTM.éÛóÃèëî≠çs = B3Chk_JRKbn.Value
                With CisComboMS
                    .CodeGet
                    ZTM.ñæç◊èëî≠çs = .Code
                End With
                ZTM.ç≈èIî≠çsòAî‘ = CisFun.Val2(B3lb_EndNO)
                
                ZTM.PNO = CisFun.Val2(iB3_PNo)
                ZTM.çﬁéøï\é¶ = B3Chk_JIS.Value              ' 2009.09.24 add
                ZTM.íIâµãÊï™ = B3Chk_Tana.Value               ' 2010.01.16 add

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
                ZTM.ëOçHíˆãÊï™ = B3Chk_Maehak.Value
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'====================================================== *** 2008/06/27 Start(Comment)
                ZTM.éÛããïiî‘ = RTrim(iB3_JHinbn)
                ZTM.î≠íçïiî‘ = RTrim(iB3_HHinbn)
'                ZTM.é¿ê—ä∑éZíl = CisFun.Val2(iB3_TKansn)                               ' 2009.02.27 del : é¿ê—ä∑éZílä÷òA
'====================================================== *** 2008/06/27 End(Comment)
'            End If                                                                     ' 2009.02.27 del : ïiñ⁄ä÷òA
'---------------------------------------------------------------------------------------' 2009.02.27 add start
                ZTM.â¡çHêÊ = RTrim(iB3_STorcd)
                'ZTM.â¡çHêÊîwåiêF = B3Btn_Hikei.BackColor                   2009/12/08
                ZTM.â¡çHêÊîwåiêF = 0
                ZTM.ìæà”êÊ = RTrim(iB3_TTorcd)
                ZTM.é©é–ïiî‘ = RTrim(iB3_SHinbn)
                ZTM.Ç©ÇÒÇŒÇÒÉTÉCÉY = B3Chk_KanBn
                ZTM.åªïiï[ÉTÉCÉY = B3Chk_GenPn
                
                ZTM.égópïiî‘1 = ""
                ZTM.égópïiî‘2 = ""
                For gCnt1 = 1 To VsGridH.Rows - 1
                    If VsGridH.TextMatrix(gCnt1, 1) <> "" And VsGridH.TextMatrix(gCnt1, 8) = "X" Then
                       If ZTM.égópïiî‘1 = "" Then
                          ZTM.égópïiî‘1 = Trim(VsGridH.TextMatrix(gCnt1, 1))
                       ElseIf ZTM.égópïiî‘2 = "" Then
                          ZTM.égópïiî‘2 = Trim(VsGridH.TextMatrix(gCnt1, 1))
                       End If
                    End If
                Next gCnt1
'---------------------------------------------------------------------------------------' 2009.02.27 add end
        End With
        ' çXêVèÓïÒÉOÉäÉbÉhï\é¶
'        Call TorInfoToGrid(wRow)
        
        .Redraw = True
    
        Call CisDB.DBTran(TransBegin)  ' ƒ◊›ªﬁ∏ºÆ›äJén
        
        If AddFlg Then
           ZTM.çÏê¨é“ = gTanto
           Call ZTMInsert
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/03/02 INSERT START
'        ' òAågçﬁóøÉ}ÉXÉ^(çÌèú)ìoò^
'           With CisDB
'                .SQL = ""
'                .SQL = .SQL & " INSERT INTO òAågçﬁóøÉ}ÉXÉ^"
'                .SQL = .SQL & " SELECT é©é–ïiî‘, '', '', ZRM.çﬁóøñº, '', ZRM.ï\é¶ê°ñ@, '', '',"
'                .SQL = .SQL & "        ZRM.ç›å…ä«óùãÊï™, ZRM.ç›å…ä«óùêîíl, ZRM.ç›å…ä«óùåWêî, ZRM.ç›å…ä«óùì˙êî, ZRM.ç›å…ä«óùî≠íçì_,"
'                .SQL = .SQL & "        '','','','','','',0,'','',"
'                .SQL = .SQL & "        ZRM.âºãÊï™,ZRM.çÏê¨ì˙,ZRM.çÏê¨é“,ZRM.çÏê¨í[ññ,ZRM.çXêVì˙,ZRM.çXêVé“,ZRM.çXêVí[ññ,"
'                .SQL = .SQL & "        ZRM.ê›íËêî,ZRM.çﬁéø,ZRM.ê°ñ@ãÊï™, ZRM.î¬å˙, ZRM.ïù, ZRM.í∑Ç≥,"
'                .SQL = .SQL & "        ZRM.çﬁéÌ, ZRM.å`èÛ, ZRM.íPà , ZRM.ä∑éZíl, ZRM.éËîzíPà , ZRM.çﬁéÌÉRÅ[Éh, ZRM.çﬁóøä«óùî‘çÜ, ZRM.èäóvó ãÊï™,"
'                .SQL = .SQL & "        CONVERT(VARCHAR(8),GETDATE(),112) + REPLACE(CONVERT(VARCHAR(8), GETDATE(),108),':',''), 2, 1"
'                .SQL = .SQL & "   FROM çﬁóøéÊà¯êÊÉ}ÉXÉ^ ZTM"
'                .SQL = .SQL & "        LEFT OUTER JOIN çﬁóøÉ}ÉXÉ^ ZRM"
'                .SQL = .SQL & "             ON ZTM.çﬁóøä«óùî‘çÜ = ZRM.çﬁóøä«óùî‘çÜ"
'                .SQL = .SQL & "  WHERE ZTM.çﬁóøä«óùî‘çÜ = " & ZTM.çﬁóøä«óùî‘çÜ & ""
'                .SQL = .SQL & "    AND ZTM.éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
'                .SQL = .SQL & "    AND ZTM.éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
'
'                .DBExec
'           End With
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/03/02 INSERT END
        Else
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/03/02 INSERT START
'        ' òAågçﬁóøÉ}ÉXÉ^(çÌèú)ìoò^
'           With CisDB
'              .SQL = ""
'              .SQL = .SQL & " INSERT INTO òAågçﬁóøÉ}ÉXÉ^"
'              .SQL = .SQL & " SELECT é©é–ïiî‘, '', '', ZRM.çﬁóøñº, '', ZRM.ï\é¶ê°ñ@, '', '',"
'              .SQL = .SQL & "        ZRM.ç›å…ä«óùãÊï™, ZRM.ç›å…ä«óùêîíl, ZRM.ç›å…ä«óùåWêî, ZRM.ç›å…ä«óùì˙êî, ZRM.ç›å…ä«óùî≠íçì_,"
'              .SQL = .SQL & "        '','','','','','',0,'','',"
'              .SQL = .SQL & "        ZRM.âºãÊï™,ZRM.çÏê¨ì˙,ZRM.çÏê¨é“,ZRM.çÏê¨í[ññ,ZRM.çXêVì˙,ZRM.çXêVé“,ZRM.çXêVí[ññ,"
'              .SQL = .SQL & "        ZRM.ê›íËêî,ZRM.çﬁéø,ZRM.ê°ñ@ãÊï™, ZRM.î¬å˙, ZRM.ïù, ZRM.í∑Ç≥,"
'              .SQL = .SQL & "        ZRM.çﬁéÌ, ZRM.å`èÛ, ZRM.íPà , ZRM.ä∑éZíl, ZRM.éËîzíPà , ZRM.çﬁéÌÉRÅ[Éh, ZRM.çﬁóøä«óùî‘çÜ, ZRM.èäóvó ãÊï™,"
'              .SQL = .SQL & "        CONVERT(VARCHAR(8),GETDATE(),112) + REPLACE(CONVERT(VARCHAR(8), GETDATE(),108),':',''), 1, 1"
'              .SQL = .SQL & "   FROM çﬁóøéÊà¯êÊÉ}ÉXÉ^ ZTM"
'              .SQL = .SQL & "        LEFT OUTER JOIN çﬁóøÉ}ÉXÉ^ ZRM"
'              .SQL = .SQL & "             ON ZTM.çﬁóøä«óùî‘çÜ = ZRM.çﬁóøä«óùî‘çÜ"
'              .SQL = .SQL & "  WHERE ZTM.çﬁóøä«óùî‘çÜ = " & ZTM.çﬁóøä«óùî‘çÜ & ""
'              .SQL = .SQL & "    AND ZTM.éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
'              .SQL = .SQL & "    AND ZTM.éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
'
'              .DBExec
'           End With
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/03/02 INSERT END
           ZTM.çXêVé“ = gTanto
           Call ZTMUpdate
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/03/02 INSERT START
'        ' òAågçﬁóøÉ}ÉXÉ^(çÌèú)ìoò^
'           With CisDB
'              .SQL = ""
'              .SQL = .SQL & " INSERT INTO òAågçﬁóøÉ}ÉXÉ^"
'              .SQL = .SQL & " SELECT é©é–ïiî‘, '', '', ZRM.çﬁóøñº, '', ZRM.ï\é¶ê°ñ@, '', '',"
'              .SQL = .SQL & "        ZRM.ç›å…ä«óùãÊï™, ZRM.ç›å…ä«óùêîíl, ZRM.ç›å…ä«óùåWêî, ZRM.ç›å…ä«óùì˙êî, ZRM.ç›å…ä«óùî≠íçì_,"
'              .SQL = .SQL & "        '','','','','','',0,'','',"
'              .SQL = .SQL & "        ZRM.âºãÊï™,ZRM.çÏê¨ì˙,ZRM.çÏê¨é“,ZRM.çÏê¨í[ññ,ZRM.çXêVì˙,ZRM.çXêVé“,ZRM.çXêVí[ññ,"
'              .SQL = .SQL & "        ZRM.ê›íËêî,ZRM.çﬁéø,ZRM.ê°ñ@ãÊï™, ZRM.î¬å˙, ZRM.ïù, ZRM.í∑Ç≥,"
'              .SQL = .SQL & "        ZRM.çﬁéÌ, ZRM.å`èÛ, ZRM.íPà , ZRM.ä∑éZíl, ZRM.éËîzíPà , ZRM.çﬁéÌÉRÅ[Éh, ZRM.çﬁóøä«óùî‘çÜ, ZRM.èäóvó ãÊï™,"
'              .SQL = .SQL & "        CONVERT(VARCHAR(8),GETDATE(),112) + REPLACE(CONVERT(VARCHAR(8), GETDATE(),108),':',''), 2, 1"
'              .SQL = .SQL & "   FROM çﬁóøéÊà¯êÊÉ}ÉXÉ^ ZTM"
'              .SQL = .SQL & "        LEFT OUTER JOIN çﬁóøÉ}ÉXÉ^ ZRM"
'              .SQL = .SQL & "             ON ZTM.çﬁóøä«óùî‘çÜ = ZRM.çﬁóøä«óùî‘çÜ"
'              .SQL = .SQL & "  WHERE ZTM.çﬁóøä«óùî‘çÜ = " & ZTM.çﬁóøä«óùî‘çÜ & ""
'              .SQL = .SQL & "    AND ZTM.éÊà¯êÊ = '" & RTrim(iB2_Torcd) & "'"
'              .SQL = .SQL & "    AND ZTM.éÛì¸ = '" & RTrim(iB2_Ukeir) & "'"
'
'              .DBExec
'           End With
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/03/02 INSERT END
        End If
            
        Call CisDB.DBTran(TransCommit)  ' ƒ◊›ªﬁ∏ºÆ›èIóπ
        
    End With
    
    DBPutTor = True
    Exit Function
DBPutTor_Err:
    CisFun.ErrorBox
    End
End Function
'+==========================================================+
'+      ÉOÉäÉbÉhèàóù                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   ê°ñ@èÓïÒ   å©èoÇµæØƒ & èâä˙èÛë‘æØƒ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'Å@∏ﬁÿØƒﬁèâä˙âªèàóù
'        .Clear
'       Call .FixedSet(å©èoÇµà íuê›íË,óÒà íuê›íË,œ∞ºﬁèàóù,
'                      æŸåÖêî,óÒî‘çÜ,å©èoÇµì‡óe,ï™óﬁ¿≤Ãﬂ)
        ReDim SCol(14)
        
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 2, 0, "No")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 8, 1, "î¬å˙")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 8, 2, "ïù")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 8, 3, "í∑Ç≥")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 0, 4, "çﬁéøä«óùî‘çÜ")
'        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 0, 5, "çﬁóøñº")
'        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 0, 6, "çﬁéÌ")
'        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 0, 7, "å`èÛ")
'        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 0, 8, "ÉRÅ[Éh")
'        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 0, 9, "íPà ")
'        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 0, 10, "ä∑éZíl")
'        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 0, 11, "éËîzíPà ")
'        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 0, 12, "ç›å…ä«óùãÊï™")

'Å@∏ﬁÿØƒﬁëIëèàóùópÅ@ëﬁî¥ÿ±
'        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 13, "Select")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 5, "Select")
'  ï\é¶ ﬂ¿∞›æØƒ
'       Call .PatanSet( ﬂ¿∞›NO,ï\é¶ÉpÉ^Å[Éì)
        Call .PatanSet(0, "0,1,2,3")
'  ï“èWì‡óeÉZÉbÉg
'       Call .EditSet(√∑ΩƒµÃﬁºﬁ™∏ƒ,∂◊—NO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
        
        SCol(1) = .FixedGet(1)  ' î¬å˙
        SCol(2) = .FixedGet(2)  ' ïù
        SCol(3) = .FixedGet(3)  ' í∑Ç≥
        SCol(4) = .FixedGet(4)  ' çﬁéøä«óùî‘çÜ
'        SCol(5) = .FixedGet(5)  ' çﬁóøñº
'        SCol(6) = .FixedGet(6)  ' çﬁéÌ
'        SCol(7) = .FixedGet(7)  ' å`èÛ
'        SCol(8) = .FixedGet(8)  ' ÉRÅ[Éh
'        SCol(9) = .FixedGet(9)  ' íPà 
'        SCol(10) = .FixedGet(10)  'ä∑éZíl
'        SCol(11) = .FixedGet(11)  'éËîzíPà 
'        SCol(12) = .FixedGet(12)  'ç›å…ä«óùãÊï™
        
    End With
    KCol = 3   ' îÒï\é¶çÄñ⁄äJén∂◊—
    
End Sub
'+----------------------------------------------------------+
'+   éËîzêÊèÓïÒ   å©èoÇµæØƒ & èâä˙èÛë‘æØƒ
'+----------------------------------------------------------+
Private Sub GridInitT()
    With CisVsGridT
'Å@∏ﬁÿØƒﬁèâä˙âªèàóù
'        .Clear
'       Call .FixedSet(å©èoÇµà íuê›íË,óÒà íuê›íË,œ∞ºﬁèàóù,
'                      æŸåÖêî,óÒî‘çÜ,å©èoÇµì‡óe,ï™óﬁ¿≤Ãﬂ)
        ReDim SColT(45)
        
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 0, 0, "No")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 20, 1, "édì¸êÊ")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 2, 2, "éÛ")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 6, 3, "îwî‘çÜ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 6, 4, "éËîz")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 10, 5, "ÉAÉhÉåÉX")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 4, 6, "î‰ó¶")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 8, 7, "SKãÊï™")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 7, 8, "ª≤∏Ÿ")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 11, 9, "é˚óeêî1")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 14, 10, "ïœçXì˙-ï÷")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 11, 11, "é˚óeêî2")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 20, 12, "å_ñÒêÊ")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 20, 13, "ÉÅÅ[ÉJÅ[")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 4, 14, "ÉÅéÛ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 22, 15, "óLå¯ä˙å¿")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 6, 16, "åWêî1")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 6, 17, "åWêî2")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 6, 18, "åWêî3")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 2, 19, "LT")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 11, 20, "€Øƒá@")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 11, 21, "€ØƒáA")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 14, 22, "ójì˙")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 14, 23, "ì˙")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 30, 24, "îıçl")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 10, 25, "é˚óeäÌ1")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 10, 26, "é˚óeäÌ2")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 10, 27, "íuèÍ")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 10, 28, "ì‡é¶î≠íç")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 10, 29, "ì`ï[éÌóﬁ")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 10, 30, "î[ïièëå`ë‘")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 10, 31, "éÛóÃèëå`ë‘")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 10, 32, "éÛóÃèëî≠çs")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 10, 33, "ñæç◊èëî≠çs")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 6, 34, "ç≈èIáÇ")

        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 4, 35, "éËîzêÊCD")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 4, 36, "ÇoáÇ")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 4, 37, "ïiñ⁄")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 4, 39, "ïiñ⁄CD")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'Å@∏ﬁÿØƒﬁëIëèàóùópÅ@ëﬁî¥ÿ±
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 38, "Select")

'---------------------------------------------------------------' 2009.02.27 add start
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 20, 40, "â¡çHêÊ")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 20, 41, "ìæà”êÊ")
'---------------------------------------------------------------' 2009.02.27 add end

'  ï\é¶ ﬂ¿∞›æØƒ
'       Call .PatanSet( ﬂ¿∞›NO,ï\é¶ÉpÉ^Å[Éì)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE START
        'Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,36")
'        Call .PatanSet(0, "0,1,2,3,37,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,36")      ' 2009.02.27 del
        Call .PatanSet(0, "0,1,2,3,37,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,36,40,41") ' 2009.02.27 add
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 UPDATE END
'  ï“èWì‡óeÉZÉbÉg
'       Call .EditSet(√∑ΩƒµÃﬁºﬁ™∏ƒ,∂◊—NO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
        
        SColT(1) = .FixedGet(1)  ' éËîzêÊÅFñºèÃ
        SColT(2) = .FixedGet(2)  ' éÛì¸
        SColT(3) = .FixedGet(3)  ' îwî‘çÜ
        SColT(4) = .FixedGet(4)  ' éËîzãÊï™ñº
        SColT(5) = .FixedGet(5)  ' ÉAÉhÉåÉX
        SColT(6) = .FixedGet(6)  ' î≠íçî‰ó¶
        SColT(7) = .FixedGet(7)  ' SKãÊï™
        SColT(8) = .FixedGet(8)  ' ÉTÉCÉNÉã
        SColT(9) = .FixedGet(9)  ' é˚óeêîÇP
        SColT(10) = .FixedGet(10)  ' ïœçXì˙-ï÷
        SColT(11) = .FixedGet(11)  ' é˚óeêîÇQ
        SColT(12) = .FixedGet(12)  ' å_ñÒêÊ
        SColT(13) = .FixedGet(13)  ' ÉÅÅ[ÉJÅ[
        SColT(14) = .FixedGet(14)  ' ÉÅÅ[ÉJÅ[éÛì¸
        SColT(15) = .FixedGet(15)  ' óLå¯ä˙å¿
        SColT(16) = .FixedGet(16)  ' åWêîÇP
        SColT(17) = .FixedGet(17)  ' åWêîÇQ
        SColT(18) = .FixedGet(18)  ' åWêîÇR
        SColT(19) = .FixedGet(19)  ' ÇkÇs
        SColT(20) = .FixedGet(20)  ' ÉçÉbÉgÇP
        SColT(21) = .FixedGet(21)  ' ÉçÉbÉgÇQ
        SColT(22) = .FixedGet(22)  ' ójì˙
        SColT(23) = .FixedGet(23)  ' ì˙
        SColT(24) = .FixedGet(24)  ' îıçl
        SColT(25) = .FixedGet(25)  ' é˚óeäÌÇP
        SColT(26) = .FixedGet(26)  ' é˚óeäÌÇQ
        SColT(27) = .FixedGet(27)  ' íuèÍ
        SColT(28) = .FixedGet(28)  ' ì‡é¶î≠íçãÊï™
        SColT(29) = .FixedGet(29)  ' ì`ï[éÌóﬁ
        SColT(30) = .FixedGet(30)  ' î[ïièëå`ë‘
        SColT(31) = .FixedGet(31)  ' éÛóÃèëå`ë‘
        SColT(32) = .FixedGet(32)  ' éÛóÃèëî≠çs
        SColT(33) = .FixedGet(33)  ' ñæç◊î≠çs
        SColT(34) = .FixedGet(34)  ' ç≈èIî≠çsî‘çÜ
        SColT(35) = .FixedGet(35)  ' éËîzêÊÇbÇc
        SColT(36) = .FixedGet(36)  ' ÇoáÇ
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
        SColT(37) = .FixedGet(37)  ' ïiñ⁄
        SColT(39) = .FixedGet(39)  ' ïiñ⁄CD
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'---------------------------------------------------------------' 2009.02.27 add start
        SColT(40) = .FixedGet(40)  ' â¡çHêÊ
        SColT(41) = .FixedGet(41)  ' ìæà”êÊ
'---------------------------------------------------------------' 2009.02.27 add end
        
    End With
    KCol = 35   ' îÒï\é¶çÄñ⁄äJén∂◊—
    
End Sub
'+----------------------------------------------------------+
'+   égópïiî‘   å©èoÇµæØƒ & èâä˙èÛë‘æØƒ
'+----------------------------------------------------------+
Private Sub GridInitH()
    With CisVsGridH
'Å@∏ﬁÿØƒﬁèâä˙âªèàóù
'        .Clear
'       Call .FixedSet(å©èoÇµà íuê›íË,óÒà íuê›íË,œ∞ºﬁèàóù,
'                      æŸåÖêî,óÒî‘çÜ,å©èoÇµì‡óe,ï™óﬁ¿≤Ãﬂ)
        ReDim SColH(11)
ReDim SColH(20)

        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 3, 0, "No")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 22, 1, "ïi  î‘")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 6, 2, "îwî‘çÜ")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 7, 3, "ÉsÉbÉ`")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 4, 4, "éÊêî")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 7, 5, "ñ¢égóp")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 7, 6, "égópó ")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 21, 7, "óLå¯ä˙å¿")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 10, 8, "Check")        ' 2009.03.03 add
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 22, 9, "ï“èWëOïiî‘")   ' 2009.03.03 add
'Å@∏ﬁÿØƒﬁëIëèàóùópÅ@ëﬁî¥ÿ±
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 4, 18, "Select")       ' 2009.03.03 upd : 8 Å® 18
'  ï\é¶ ﬂ¿∞›æØƒ
'       Call .PatanSet( ﬂ¿∞›NO,ï\é¶ÉpÉ^Å[Éì)
        Call .PatanSet(0, "0,1")
'       Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,18")             ' 2009.03.03 ééå±óp
'  ï“èWì‡óeÉZÉbÉg
'       Call .EditSet(√∑ΩƒµÃﬁºﬁ™∏ƒ,∂◊—NO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
        
        SColH(1) = .FixedGet(1)     ' ïiî‘
        SColH(2) = .FixedGet(2)     ' îwî‘çÜ
        SColH(3) = .FixedGet(3)     ' ÉsÉbÉ`
        SColH(4) = .FixedGet(4)     ' éÊêî
        SColH(5) = .FixedGet(5)     ' égópó 
        SColH(6) = .FixedGet(6)     ' ñ¢égóp        ' 2009.03.03 / ÉRÉÅÉìÉgïœçX ' óLå¯ä˙å¿
        SColH(7) = .FixedGet(7)     ' égópó         ' 2009.03.03 / ÉRÉÅÉìÉgïœçX ' ï“èWëOïiî‘
        SColH(8) = .FixedGet(8)     ' Check         ' 2009.03.03 / ÉRÉÅÉìÉgïœçX ' äJénì˙
        SColH(9) = .FixedGet(9)     ' ï“èWëOïiî‘    ' 2009.03.03 / ÉRÉÅÉìÉgïœçX ' èIóπì˙
        SColH(18) = .FixedGet(18)   ' Select        ' 2009.03.03 add
    End With
    KCol = 8   ' îÒï\é¶çÄñ⁄äJén∂◊—
    
End Sub
'****************************************************************************************'
'*      ÉOÉäÉbÉhì‡óeÉZÉbÉg(ÉTÉCÉYèÓïÒ)
'****************************************************************************************'
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
    Do Until Not ZRM_RDSTS
        ID = ID + 1
            
        Call ZaiInfoToGrid(ID)
        
        Call ZRMReadNext
    Loop
    With vsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'+---------------------------------------------------------------------------------------+
'+  çﬁóøèÓïÒÇÊÇËÉOÉäÉbÉhÇ÷ÉZÉbÉg
'+---------------------------------------------------------------------------------------+
Private Sub ZaiInfoToGrid(SetRow As Long)
    With vsGrid1
        For mIx = 0 To .Cols - 1
            .TextMatrix(SetRow, mIx) = ""
        Next mIx
        
        ' No
        .TextMatrix(SetRow, 0) = SetRow
        .TextMatrix(SetRow, 1) = CisFun.RSetFld(ZRM.î¬å˙, 7, "###0.00")
        .TextMatrix(SetRow, 2) = CisFun.RSetFld(ZRM.ïù, 7, "###0.00")
        .TextMatrix(SetRow, 3) = CisFun.RSetFld(ZRM.í∑Ç≥, 7, "###0.00")
        .TextMatrix(SetRow, 4) = ZRM.çﬁóøä«óùî‘çÜ
        .TextMatrix(SetRow, 5) = ZRM.çﬁóøñº
        .TextMatrix(SetRow, 6) = ZRM.çﬁéÌ
        .TextMatrix(SetRow, 7) = ZRM.å`èÛ
        .TextMatrix(SetRow, 8) = ZRM.çﬁéÌÉRÅ[Éh
        .TextMatrix(SetRow, 9) = ZRM.íPà 
        .TextMatrix(SetRow, 10) = ZRM.ä∑éZíl
        .TextMatrix(SetRow, 11) = ZRM.éËîzíPà 
        .TextMatrix(SetRow, 12) = ZRM.ç›å…ä«óùãÊï™
        
        If Rv_Call3 <> "" Then
            If RV_Itatu = ZRM.î¬å˙ And _
               RV_Width = ZRM.ïù And _
               RV_Long = ZRM.í∑Ç≥ Then
                wsROW = SetRow
            End If
        End If
    
    End With
End Sub
'+---------------------------------------------------------------------------------------+
'+  çﬁóøèÓïÒÉOÉäÉbÉhÇÊÇËì¸óÕàÊÇ÷ÉZÉbÉg
'+---------------------------------------------------------------------------------------+
Private Sub GridToText()
    
    With vsGrid1
        If Trim(.TextMatrix(.Row, SCol(4))) = "" Then Exit Sub
        
    ' * H2
        iH2_Itatu.Tag = CisFun.RSetFld(.TextMatrix(.Row, SCol(1)), 7, "###0.00")
        iH2_Width.Tag = CisFun.RSetFld(.TextMatrix(.Row, SCol(2)), 7, "###0.00")
        iH2_Long.Tag = CisFun.RSetFld(.TextMatrix(.Row, SCol(3)), 7, "###0.00")
        iH2_Itatu = CisFun.RSetFld(.TextMatrix(.Row, SCol(1)), 7, "###0.00")
        iH2_Width = CisFun.RSetFld(.TextMatrix(.Row, SCol(2)), 7, "###0.00")
        iH2_Long = CisFun.RSetFld(.TextMatrix(.Row, SCol(3)), 7, "###0.00")
        
    
        wKanri = CisFun.Val2(.TextMatrix(.Row, SCol(4)))
        gSL_Select = "SELECT * FROM çﬁóøÉ}ÉXÉ^ "
        gSL_Select = gSL_Select & " WHERE çﬁéø = '" & RTrim(iH1_Zaist) & "'"
        gSL_Select = gSL_Select & "   AND çﬁóøä«óùî‘çÜ = " & wKanri
        Call ZRMRead(gSL_Select, 1)
    
    ' * B1
        iB1_Name = RTrim(ZRM.çﬁóøñº)
        iB1_Zaisy = RTrim(ZRM.çﬁéÌ)
        Call iB1_Zaisy_LostFocus
        iB1_Keijo = RTrim(ZRM.å`èÛ)
        Call iB1_Keijo_LostFocus
        iB1_Code = RTrim(ZRM.çﬁéÌÉRÅ[Éh)
        iB1_Tani = RTrim(ZRM.íPà )
        Call iB1_Tani_LostFocus
        iB1_Kansn = CisFun.RSetFld(ZRM.ä∑éZíl, 10, "###0.000")
        iB1_TTani = RTrim(ZRM.éËîzíPà )
        Call iB1_TTani_LostFocus
        B1Chk_JSKbn.Value = ZRM.é¿êîãÊï™
        B1Chk_ZKKbn.Value = ZRM.ç›å…ãÊï™
        
        Call CisComboZK.ComboSet
        With CisComboZK
            .Code = ZRM.ç›å…ä«óùãÊï™
            .TextGet
        End With
        iB1_Zaiko = CisFun.RSetFld(ZRM.ç›å…ä«óùåWêî, 0, iB1_Zaiko.cFormat)
        iB1_ZaikoDay = CisFun.RSetFld(ZRM.ç›å…ä«óùì˙êî, 0, iB1_ZaikoDay.cFormat)
        iB1_ZaikoL = CisFun.RSetFld(ZRM.ç›å…ä«óùî≠íçì_, 0, iB1_ZaikoL.cFormat)
        
' ===================================================== *** 2008/06/27 Start(Comment)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
'        iB1_JHinbn = RTrim(ZRM.éÛããïiî‘)
'        iB1_HHinbn = RTrim(ZRM.î≠íçïiî‘)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
' ===================================================== *** 2008/06/27 End(Comment)

        B1Op_Syoyo(ZRM.èäóvó ãÊï™).Value = True     ' 2009.02.27 add
    
    End With

' éËîzêÊèÓïÒ(ÉOÉäÉbÉh)
    RowColFlg = True
    Call GridSetT
    RowColFlg = False

' égópïiî‘èÓïÒ(ÉOÉäÉbÉh)
    Call GridSetH

End Sub
'+---------------------------------------+
'+  ÉOÉäÉbÉh (Click) ÉTÉCÉYèÓïÒ
'+---------------------------------------+
Private Sub VSGrid1_ClickClick()
    
    With vsGrid1
        If ProcHB = "H2" Then
            GridBack = ProcHB
'            Call HeadBodyClear("H2")
            iH2_Itatu = ""
            iH2_Width = ""
            iH2_Long = ""
            Call HeadBodyClear("B1")
            
            If SyoriKB <> "ADD" Then
                If Trim(.TextMatrix(.Row, 4)) <> "" Then
                    RowColFlg = True
                    Call CisVsGrid.vsColor(è¡ãéëSïî)
                    Call CisVsGrid.vsColor(å¬ï )
                    
                    RowColFlg = False
                    
                    Call GridToText
                   
                End If
            End If
        End If
    End With
End Sub
'+---------------------------------------+
'+  ÉOÉäÉbÉh (DblClick) ÉTÉCÉYèÓïÒ
'+---------------------------------------+
Private Sub vsGrid1_DblClick()
    With vsGrid1
        If ProcHB = "H2" Then
            If Trim(.TextMatrix(.Row, 5)) <> "" Then
                If SyoriKB2 <> "DEL" Then
                    RowColFlg = True
                    Call CisVsGrid.vsColor(è¡ãéëSïî)
                    Call CisVsGrid.vsColor(å¬ï )
                    RowColFlg = False
                    
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/05/15 UPDATE START
                    'SyoriKB2 = "MNT"
                    'Call SyoriName2(SyoriKB2)
                    
                    Call GridToText
                    SyoriKB = "MNT"
                    Call SyoriName(SyoriKB)
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/05/15 UPDATE END
                    ProcHB = "B1"
                    Call DispChange(ProcHB)
                    ' ∏ﬁÿØƒﬁì‡óeï\é¶å„ÅA√∑Ωƒì¸óÕ
'                    Call CisVsGridS.EditVisible("C")
                    
                    iB1_Name.SetFocus
                End If
            End If
        End If
    End With
End Sub
'+---------------------------------------+
'+  ÉOÉäÉbÉh (GotFocus) ÉTÉCÉYèÓïÒ
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
'+  ÉOÉäÉbÉh (GotFocus) ÉTÉCÉYèÓïÒ
'+---------------------------------------+
Private Sub vsGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
       Call VSGrid1_ClickClick
       Call vsGrid1_DblClick
    End If
End Sub
'+---------------------------------------+
'+  ÉOÉäÉbÉh (RowColChange) ÉTÉCÉYèÓïÒ
'+---------------------------------------+
Private Sub vsGrid1_RowColChange()
    If Not FormAct Then Exit Sub
    If RowColFlg Then Exit Sub
    Call VSGrid1_ClickClick
End Sub
'****************************************************************************************'
'*      ÉOÉäÉbÉhì‡óeÉZÉbÉg(éÊà¯êÊèÓïÒ)
'****************************************************************************************'

Private Sub GridSetT()
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

' éËîzêÊÅ@ÉZÉbÉg
    gSL_Select = "SELECT * FROM çﬁóøéÊà¯êÊÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " WHERE çﬁóøä«óùî‘çÜ = " & wKanri
    Call ZTMRead(gSL_Select, , 1)

' íäèoèàóù èâä˙æØƒ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' íäèoåèêî(çáåv)ëﬁî
' ∏ﬁÿØƒﬁçsçƒæØƒ & óÒà íuéÊìæ
    With VsGridT
       .Redraw = False
        
        If Ttl_Kensu + 1 < CisVsGridT.DispMax Then
            .Rows = CisVsGridT.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not ZTM_RDSTS
        ID = ID + 1
            
        Call TorInfoToGrid(ID)
        
        vsGrid1.Redraw = True
        DoEvents
        Call ZTMReadNext(1)
    Loop
' ÉNÉçÅ[ÉY
    Call ZTMClose
       
    With VsGridT
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'+----------------------------------------------------------------------------------+
'+      çﬁóøéÊà¯êÊèÓïÒÇÊÇËÉOÉäÉbÉhÇ÷ÉZÉbÉg
'+----------------------------------------------------------------------------------+
Private Sub TorInfoToGrid(SetRow As Long)
    With VsGridT
        For mIx = 0 To .Cols - 1
            .TextMatrix(SetRow, mIx) = ""
        Next mIx
        
        ' No
        .TextMatrix(SetRow, 0) = SetRow
        'éËîzêÊ
        If TorNmGet(ZTM.éÊà¯êÊ, 9) Then
           .TextMatrix(SetRow, SColT(1)) = ZTM.éÊà¯êÊ & ":" & TRM_RName
        Else
           .TextMatrix(SetRow, SColT(1)) = ZTM.éÊà¯êÊ
        End If
        .TextMatrix(SetRow, SColT(35)) = ZTM.éÊà¯êÊ
        ' éÛì¸
        .TextMatrix(SetRow, SColT(2)) = ZTM.éÛì¸
        ' îwî‘çÜ
        .TextMatrix(SetRow, SColT(3)) = ZTM.îwî‘çÜ
        ' éËîzãÊï™
'        If ZTM.ïiñ⁄ = 1 Then       '======== 2008/02/06 ADD            ' 2009.02.27 del : ïiñ⁄ä÷òA
            With CisComboTH
                .Code = ZTM.éËîzãÊï™
                .NameGet
                VsGridT.TextMatrix(SetRow, SColT(4)) = .CombText
            End With
'        End If                              '======== 2008/02/06 ADD   ' 2009.02.27 del : ïiñ⁄ä÷òA
        
        ' ÉAÉhÉåÉX
        .TextMatrix(SetRow, SColT(5)) = ZTM.ÉAÉhÉåÉX
        ' î‰ó¶
        If ZTM.î≠íçî‰ó¶ <> 0 Then
           .TextMatrix(SetRow, SColT(6)) = ZTM.î≠íçî‰ó¶ & "%"
        End If
        ' SK
        If ZTM.SKãÊï™ = 0 Then
           .TextMatrix(SetRow, SColT(7)) = "Ç©ÇÒÇŒÇÒ"
        Else
           .TextMatrix(SetRow, SColT(7)) = "éwé¶"
        End If
        ' ÉTÉCÉNÉã
        If Trim(ZTM.ÉTÉCÉNÉã) <> "" Then
            .TextMatrix(SetRow, SColT(8)) = Mid(ZTM.ÉTÉCÉNÉã, 1, 1) & "-" & _
                                     Mid(ZTM.ÉTÉCÉNÉã, 2, 2) & "-" & _
                                     Mid(ZTM.ÉTÉCÉNÉã, 4, 2)
        End If
        ' é˚óeêî1
        If ZTM.é˚óeêî1 <> 0 Then
            .TextMatrix(SetRow, SColT(9)) = Format(ZTM.é˚óeêî1, "#,##0.000")
        End If
        ' ïœçXì˙Å[ï÷
        If Trim(ZTM.ïœçXì˙) <> "" Then
           .TextMatrix(SetRow, SColT(10)) = Mid(ZTM.ïœçXì˙, 1, 4) & "/" & _
                                     Mid(ZTM.ïœçXì˙, 5, 2) & "/" & _
                                     Mid(ZTM.ïœçXì˙, 7, 2) & "-" & _
                                     Format(CisFun.Val2(ZTM.ïœçXï÷), "00")
        End If
        ' é˚óeêî2
        If ZTM.é˚óeêî2 <> 0 Then
            .TextMatrix(SetRow, SColT(11)) = Format(ZTM.é˚óeêî2, "#,##0.000")
        End If
        ' å_ñÒêÊ
        If TorNmGet(ZTM.å_ñÒêÊ, 5) Then
           .TextMatrix(SetRow, SColT(12)) = ZTM.å_ñÒêÊ & ":" & TRM_RName
        Else
           .TextMatrix(SetRow, SColT(12)) = ZTM.å_ñÒêÊ
        End If
        ' î[ì¸êÊ
        If TorNmGet(ZTM.ÉÅÅ[ÉJÅ[, 8) Then
           .TextMatrix(SetRow, SColT(13)) = ZTM.ÉÅÅ[ÉJÅ[ & ":" & TRM_RName
        Else
           .TextMatrix(SetRow, SColT(13)) = ZTM.ÉÅÅ[ÉJÅ[
        End If
        ' î[ì¸êÊéÛì¸
        .TextMatrix(SetRow, SColT(14)) = ZTM.ÉÅÅ[ÉJÅ[éÛì¸
        ' óLå¯ä˙å¿
        gStr = ""
        If Trim(ZTM.óLå¯ä˙å¿äJén) <> "" Then
            gStr = Mid(ZTM.óLå¯ä˙å¿äJén, 1, 4) & "/" & _
                   Mid(ZTM.óLå¯ä˙å¿äJén, 5, 2) & "/" & _
                   Mid(ZTM.óLå¯ä˙å¿äJén, 7, 2)
        End If
        If Trim(ZTM.óLå¯ä˙å¿èIóπ) <> "" Then
           If gStr = "" Then gStr = Space(10)
           gStr = gStr & "Å`"
           gStr = gStr & Mid(ZTM.óLå¯ä˙å¿èIóπ, 1, 4) & "/" & _
                         Mid(ZTM.óLå¯ä˙å¿èIóπ, 5, 2) & "/" & _
                         Mid(ZTM.óLå¯ä˙å¿èIóπ, 7, 2)
        Else
           If Trim(ZTM.óLå¯ä˙å¿äJén) <> "" Then
              gStr = gStr & "Å`" & Space(10)
           End If
        End If
        .TextMatrix(SetRow, SColT(15)) = gStr
        
        ' åWêî1
        If ZTM.åWêî1 <> 0 Then
           .TextMatrix(SetRow, SColT(16)) = ZTM.åWêî1
        End If
        ' åWêî2
        If ZTM.åWêî2 <> 0 Then
           .TextMatrix(SetRow, SColT(17)) = ZTM.åWêî2
        End If
        ' åWêî3
        If ZTM.åWêî3 <> 0 Then
           .TextMatrix(SetRow, SColT(18)) = ZTM.åWêî3
        End If
        ' LT
        If ZTM.LT <> 0 Then
           .TextMatrix(SetRow, SColT(19)) = ZTM.LT
        End If
        ' ÉçÉbÉgá@
        If ZTM.ÉçÉbÉg1 <> 0 Then
           .TextMatrix(SetRow, SColT(20)) = Format(ZTM.ÉçÉbÉg1, iB3_Lot1.cFormat)
        End If
        ' ÉçÉbÉgáA
        If ZTM.ÉçÉbÉg2 <> 0 Then
           .TextMatrix(SetRow, SColT(21)) = Format(ZTM.ÉçÉbÉg2, iB3_Lot2.cFormat)
        End If
        ' ójì˙
        gStr = "": mStr = "•"
        If ZTM.Week1 <> 0 Then
           gStr = B3Chk_Week(0).Caption
        Else
           gStr = "Å@"
        End If
        If ZTM.Week2 <> 0 Then
           gStr = gStr & mStr & B3Chk_Week(1).Caption
        Else
           gStr = gStr & mStr & "Å@"
        End If
        If ZTM.Week3 <> 0 Then
           gStr = gStr & mStr & B3Chk_Week(2).Caption
        Else
           gStr = gStr & mStr & "Å@"
        End If
        If ZTM.Week4 <> 0 Then
           gStr = gStr & mStr & B3Chk_Week(3).Caption
        Else
           gStr = gStr & mStr & "Å@"
        End If
        If ZTM.Week5 <> 0 Then
           gStr = gStr & mStr & B3Chk_Week(4).Caption
        Else
           gStr = gStr & mStr & "Å@"
        End If
        If gStr <> "" Then
            gStr = Replace(gStr, "Å@•Å@•Å@•Å@•Å@", "")
           .TextMatrix(SetRow, SColT(22)) = gStr
        End If
        ' ì˙
        gStr = ""
        If ZTM.Day1 <> 0 Then
           gStr = CisFun.RSetFld(ZTM.Day1, 0, iB3_Day(0).cFormat)
        End If
        If ZTM.Day2 <> 0 Then
           If gStr <> "" Then gStr = gStr & "•"
           gStr = gStr & CisFun.RSetFld(ZTM.Day2, 0, iB3_Day(1).cFormat)
        End If
        If ZTM.Day3 <> 0 Then
           If gStr <> "" Then gStr = gStr & "•"
           gStr = gStr & CisFun.RSetFld(ZTM.Day3, 0, iB3_Day(2).cFormat)
        End If
        If ZTM.Day4 <> 0 Then
           If gStr <> "" Then gStr = gStr & "•"
           gStr = gStr & CisFun.RSetFld(ZTM.Day4, 0, iB3_Day(3).cFormat)
        End If
        If ZTM.Day5 <> 0 Then
           If gStr <> "" Then gStr = gStr & "•"
           gStr = gStr & CisFun.RSetFld(ZTM.Day5, 0, iB3_Day(4).cFormat)
        End If
        If gStr <> "" Then
           .TextMatrix(SetRow, SColT(23)) = gStr
        End If
        ' îıçl
         .TextMatrix(SetRow, SColT(24)) = ZTM.îıçl
        ' é˚óeäÌ1
        .TextMatrix(SetRow, SColT(25)) = ZTM.é˚óeäÌ1
        ' é˚óeäÌ2
        .TextMatrix(SetRow, SColT(26)) = ZTM.é˚óeäÌ2
        ' íuèÍ
        '================================================- 2009/12/08 Start
        '.TextMatrix(SetRow, SColT(27)) = ZTM.íuèÍ
        If RTrim(ZTM.íuèÍ) <> "" Then
           .TextMatrix(SetRow, SColT(27)) = Mid(ZTM.íuèÍ, 1, 2) & Space(2 - Len(Mid(ZTM.íuèÍ, 1, 2))) & "-" & Mid(ZTM.íuèÍ, 3, 2) & Space(2 - Len(Mid(ZTM.íuèÍ, 3, 2))) & "-" & Mid(ZTM.íuèÍ, 5, 2) & Space(2 - Len(Mid(ZTM.íuèÍ, 5, 2))) & "-" & Mid(ZTM.íuèÍ, 7, 2) & Space(2 - Len(Mid(ZTM.íuèÍ, 7, 2)))
        Else
           .TextMatrix(SetRow, SColT(27)) = ""
        End If
        '================================================- 2009/12/08 End
        ' ì‡é¶î≠íç
        If ZTM.ì‡é¶î≠íçãÊï™ = 0 Then
            .TextMatrix(SetRow, SColT(28)) = "Ç∑ÇÈ"
        Else
            .TextMatrix(SetRow, SColT(28)) = "ÇµÇ»Ç¢"
        End If
    
        ' ì`ï[éÌóﬁ
        If MeisyoGet("ì`ï[éÌóﬁî≠íç", ZTM.ì`ï[éÌóﬁ) Then
            .TextMatrix(SetRow, SColT(29)) = SYM_Meisyo
        End If
        ' î[ïièëå`ë‘
        If MeisyoGet("î[ïièëå`ë‘î≠íç", ZTM.î[ïièëå`ë‘) Then
            .TextMatrix(SetRow, SColT(30)) = SYM_Meisyo
        End If
        ' éÛóÃèëå`ë‘
        If MeisyoGet("éÛóÃèëå`ë‘", ZTM.éÛóÃèëå`ë‘) Then
            .TextMatrix(SetRow, SColT(31)) = SYM_Meisyo
        End If
        ' éÛóÃèëî≠çs
        If ZTM.éÛóÃèëî≠çs = 0 Then
            .TextMatrix(SetRow, SColT(32)) = "î≠íçéûî≠çs"
        Else
            .TextMatrix(SetRow, SColT(32)) = "éÛóÃéûî≠çs"
        End If
        ' ñæç◊î≠çs
        If MeisyoGet("ñæç◊èëî≠çs", ZTM.ñæç◊èëî≠çs) Then
            .TextMatrix(SetRow, SColT(33)) = SYM_Meisyo
        End If
    
        ' PNO
        If ZTM.PNO <> 0 Then
           .TextMatrix(SetRow, SColT(36)) = ZTM.PNO
        End If
'---------------------------------------------------------------------------------------' 2009.02.27 add start
        ' â¡çHêÊ
        If TorNmGet(ZTM.â¡çHêÊ, 7) Then
           .TextMatrix(SetRow, SColT(40)) = ZTM.â¡çHêÊ & ":" & TRM_RName
        Else
           .TextMatrix(SetRow, SColT(40)) = ZTM.â¡çHêÊ
        End If
        ' ìæà”êÊ
        If TorNmGet(ZTM.ìæà”êÊ, 2) Then
           .TextMatrix(SetRow, SColT(41)) = ZTM.ìæà”êÊ & ":" & TRM_RName
        Else
           .TextMatrix(SetRow, SColT(41)) = ZTM.ìæà”êÊ
        End If
'---------------------------------------------------------------------------------------' 2009.02.27 add end
        
'---------------------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA start
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
'        ' ïiñ⁄
'        If ZTM.ïiñ⁄ = 0 Then
'           .TextMatrix(SetRow, SColT(37)) = "éÛíç"
'        Else
'           .TextMatrix(SetRow, SColT(37)) = "î≠íç"
'        End If
'        .TextMatrix(SetRow, SColT(39)) = ZTM.ïiñ⁄
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'---------------------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA end
'---------------------------------------------------------------------------------------' 2009.02.27 add : ïiñ⁄ä÷òA start
        .TextMatrix(SetRow, SColT(37)) = "î≠íç"
        .TextMatrix(SetRow, SColT(39)) = 1
'---------------------------------------------------------------------------------------' 2009.02.27 add : ïiñ⁄ä÷òA end
    
    End With
End Sub
'*************************************************
'*      éÊà¯êÊèÓïÒÉOÉäÉbÉhÇÊÇËì¸óÕàÊÇ÷ÉZÉbÉg        *
'*************************************************
Private Sub GridToTextT()
    mIx = CisFun.Val2(iB2_Torcd.Tag)
    
    Call CisComboTH.ComboSet
'    Call CisComboZK.ComboSet
    
    If Trim(VsGridT.TextMatrix(mIx, SColT(35))) = "" Then Exit Sub
    
    gSL_Select = "SELECT * FROM çﬁóøéÊà¯êÊÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " WHERE çﬁóøä«óùî‘çÜ = " & wKanri
    gSL_Select = gSL_Select & "   AND éÊà¯êÊ = '" & RTrim(VsGridT.TextMatrix(mIx, SColT(35))) & "'"
    gSL_Select = gSL_Select & "   AND éÛì¸ = '" & RTrim(VsGridT.TextMatrix(mIx, SColT(2))) & "'"
'    gSL_Select = gSL_Select & "   AND ïiñ⁄ = " & CisFun.Val2(VsGridT.TextMatrix(mIx, SColT(39)))       ' 2009.02.27 del
    
    If ZTMRead(gSL_Select, 1) Then
        
        ' éÊà¯êÊ
        iB2_Torcd = ZTM.éÊà¯êÊ
        Call iB2_Torcd_LostFocus
        ' éÛì¸
        iB2_Ukeir = ZTM.éÛì¸
        ' îwî‘çÜ
        iB3_Seban = ZTM.îwî‘çÜ
        ' SK
        If ZTM.SKãÊï™ = 0 Then
           B3Op_SKKbn(0).Value = True
        Else
           B3Op_SKKbn(1).Value = True
        End If
        ' ÉTÉCÉNÉã
        gStr = ZTM.ÉTÉCÉNÉã
        If gStr <> "" Then
           gStr = Replace(gStr, "-", "")
           iB3_Cycle1 = Mid(gStr, 1, 1)
           iB3_Cycle2 = Mid(gStr, 2, 2)
           iB3_Cycle3 = Mid(gStr, 4, 2)
        End If
        ' é˚óeêî1
        If ZTM.é˚óeêî1 <> 0 Then: iB3_Syuyo1 = CisFun.RSetFld(ZTM.é˚óeêî1, 0, iB3_Syuyo1.cFormat)
        ' é˚óeäÌ1
        iB3_Hako1 = ZTM.é˚óeäÌ1
        ' ïœçXì˙Å[ï÷
        iB3_CDate = Trim(ZTM.ïœçXì˙)
        If CisFun.Val2(ZTM.ïœçXï÷) <> 0 Then: iB3_CBin = CisFun.RSetFld(ZTM.ïœçXï÷, 0, iB3_CBin.cFormat)
        ' é˚óeêî2
        If ZTM.é˚óeêî2 <> 0 Then: iB3_Syuyo2 = CisFun.RSetFld(ZTM.é˚óeêî2, 0, iB3_Syuyo2.cFormat)
        ' é˚óeäÌ2
        iB3_Hako2 = ZTM.é˚óeäÌ2
        ' ç≈èIÇmÇn
        If ZTM.ç≈èIî≠çsòAî‘ <> 0 Then: B3lb_EndNO = CisFun.RSetFld(ZTM.ç≈èIî≠çsòAî‘, 3, "000")
        
        ' íuèÍ
        '============================================== 2009/12/08 Start
        'iB3_Okiba = ZTM.íuèÍ
        iB3_Okiba1 = Mid(ZTM.íuèÍ, 1, 2)
        iB3_Okiba2 = Mid(ZTM.íuèÍ, 3, 2)
        iB3_Okiba3 = Mid(ZTM.íuèÍ, 5, 2)
        iB3_Okiba4 = Mid(ZTM.íuèÍ, 7, 2)
        '============================================== 2009/12/08 End
        ' å_ñÒêÊ
        iB3_KTorcd = ZTM.å_ñÒêÊ
        Call iB3_KTorcd_LostFocus
        ' ÉAÉhÉåÉX
        iB3_ADD = ZTM.ÉAÉhÉåÉX
        ' óLå¯ä˙å¿(äJén)
        iB3_SDate = ZTM.óLå¯ä˙å¿äJén
        ' óLå¯ä˙å¿(èIóπ)
        iB3_EDate = ZTM.óLå¯ä˙å¿èIóπ
        
        ' éËîzãÊï™
        Call CisComboTH.ComboSet
        With CisComboTH
            .Code = ZTM.éËîzãÊï™
            .TextGet
            If .Code = "2" Then
               B3Comb_THKbn.Visible = False
            Else
               B3Comb_THKbn.Visible = True
            End If
        End With
        
        ' ÉÅÅ[ÉJÅ[
        iB3_MTorcd = ZTM.ÉÅÅ[ÉJÅ[
        Call iB3_MTorcd_LostFocus
        ' î[ì¸êÊéÛì¸
        iB3_MUkeir = ZTM.ÉÅÅ[ÉJÅ[éÛì¸
        ' î‰ó¶
        If ZTM.î≠íçî‰ó¶ <> 0 Then: iB3_HRitu = CisFun.RSetFld(ZTM.î≠íçî‰ó¶, 0, iB3_HRitu.cFormat)
        ' ì‡é¶î≠íç
        B3Chk_NHKbn.Value = ZTM.ì‡é¶î≠íçãÊï™
        ' îıçl
        iB3_Biko = ZTM.îıçl

        ' åWêî1
        If ZTM.åWêî1 <> 0 Then: iB3_Keisu1 = CisFun.RSetFld(ZTM.åWêî1, 0, iB3_Keisu1.cFormat)
        ' åWêî2
        If ZTM.åWêî2 <> 0 Then: iB3_Keisu2 = CisFun.RSetFld(ZTM.åWêî2, 0, iB3_Keisu2.cFormat)
        ' åWêî3
        If ZTM.åWêî3 <> 0 Then: iB3_Keisu3 = CisFun.RSetFld(ZTM.åWêî3, 0, iB3_Keisu3.cFormat)
        ' LT
        If ZTM.LT <> 0 Then: iB3_LT = CisFun.RSetFld(ZTM.LT, 0, iB3_LT.cFormat)
        ' ÉçÉbÉgá@
        If ZTM.ÉçÉbÉg1 <> 0 Then
            iB3_Lot1 = CisFun.RSetFld(ZTM.ÉçÉbÉg1, 11, iB3_Lot1.cFormat)
        End If
        ' ÉçÉbÉgáA
        If ZTM.ÉçÉbÉg2 <> 0 Then
            iB3_Lot2 = CisFun.RSetFld(ZTM.ÉçÉbÉg2, 11, iB3_Lot2.cFormat)
        End If
        ' ójì˙
        B3Chk_Week(0).Value = ZTM.Week1
        B3Chk_Week(1).Value = ZTM.Week2
        B3Chk_Week(2).Value = ZTM.Week3
        B3Chk_Week(3).Value = ZTM.Week4
        B3Chk_Week(4).Value = ZTM.Week5
        ' ì˙
        If ZTM.Day1 <> 0 Then: iB3_Day(0) = CisFun.RSetFld(ZTM.Day1, 2, "##")
        If ZTM.Day2 <> 0 Then: iB3_Day(1) = CisFun.RSetFld(ZTM.Day2, 2, "##")
        If ZTM.Day3 <> 0 Then: iB3_Day(2) = CisFun.RSetFld(ZTM.Day3, 2, "##")
        If ZTM.Day4 <> 0 Then: iB3_Day(3) = CisFun.RSetFld(ZTM.Day4, 2, "##")
        If ZTM.Day5 <> 0 Then: iB3_Day(4) = CisFun.RSetFld(ZTM.Day5, 2, "##")
        
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
'---------------------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA start
'        If ZTM.ïiñ⁄ = 0 Then
'           B2Op_Hinmoku(0).Value = True
'        Else
'           B2Op_Hinmoku(1).Value = True
'        End If
'---------------------------------------------------------------------------------------' 2009.02.27 del : ïiñ⁄ä÷òA end

        B2Op_Hinmoku(1).Value = True                                                    ' 2009.02.27 add : ïiñ⁄ä÷òA
        If B2Op_Hinmoku(0).Value Then
            iB3_DensyJ = ZTM.ì`ï[éÌóﬁ
            Call iB3_DensyJ_LostFocus
            B3Chk_PrtkbNJ.Value = ZTM.î[ïièëî≠çsãÊï™
            B3Chk_PrtkbSJ.Value = ZTM.éxããèëî≠çsãÊï™
            B3Chk_PrtkbJJ.Value = ZTM.éÛóÃèëî≠çsãÊï™
            B3Chk_JRKbnJ.Value = ZTM.éÛóÃèëî≠çs
            iB3_PNoJ = CisFun.RSetFld(ZTM.PNO, 2, "##")
        Else
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
            ' ì`ï[éÌóﬁ
            iB3_Densy = ZTM.ì`ï[éÌóﬁ
            Call iB3_Densy_LostFocus
'''2007/12/11DEL        ' î[ïièëå`ë‘
'''2007/12/11DEL        Call CisComboNH.ComboSet
'''2007/12/11DEL        With CisComboNH
'''2007/12/11DEL            .Code = ZTM.î[ïièëå`ë‘
'''2007/12/11DEL            .TextGet
'''2007/12/11DEL        End With
'''2007/12/11DEL        ' î[ïièëå`ë‘
'''2007/12/11DEL        Call CisComboJR.ComboSet
'''2007/12/11DEL        With CisComboJR
'''2007/12/11DEL            .Code = ZTM.éÛóÃèëå`ë‘
'''2007/12/11DEL            .TextGet
'''2007/12/11DEL        End With
            ' éÛóÃèëî≠çs
            
            '============================================= ***** 2008/06/11 Start
            'B3Chk_PrtkbN.Value = ZTM.î[ïièëî≠çsãÊï™
            If ZTM.éËîzãÊï™ <> 4 Then
               B3Chk_PrtkbN.Value = ZTM.î[ïièëî≠çsãÊï™
            Else
               B3Chk_PrtkbSN.Value = ZTM.î[ïièëî≠çsãÊï™
            End If
            '============================================= ***** 2008/06/11 End
'---------------------------------------------------------------------------------------' 2009.02.27 del : íºëóì`ï[î≠çsãÊï™ä÷òA start
'            '============================================= 2009/01/15 Start
'            If ZTM.éËîzãÊï™ = 1 Then
'               B3Chk_PrtkbSN.Value = ZTM.íºëóì`ï[î≠çsãÊï™
'            End If
'            '============================================= 2009/01/15 End
'---------------------------------------------------------------------------------------' 2009.02.27 del : íºëóì`ï[î≠çsãÊï™ä÷òA end
            
            B3Chk_PrtkbS.Value = ZTM.éxããèëî≠çsãÊï™
            B3Chk_PrtkbJ.Value = ZTM.éÛóÃèëî≠çsãÊï™
            
            B3Chk_JRKbn.Value = ZTM.éÛóÃèëî≠çs
            ' ñæç◊èëî≠çs
            Call CisComboMS.ComboSet
            With CisComboMS
                .Code = ZTM.ñæç◊èëî≠çs
                .TextGet
            End With
        
            ' ÇoáÇ
            iB3_PNo = CisFun.RSetFld(ZTM.PNO, 2, "##")
            ' çﬁéøï\é¶
            If ZTM.çﬁéøï\é¶ <> 1 Then ZTM.çﬁéøï\é¶ = 0      ' 2009.09.24 add
            B3Chk_JIS.Value = ZTM.çﬁéøï\é¶                  ' 2009.09.24 add
            B3Chk_Tana.Value = ZTM.íIâµãÊï™                 ' 2010.01.16 add

' ===================================================== *** 2008/06/27 Start
            iB3_JHinbn = RTrim(ZTM.éÛããïiî‘)
            iB3_HHinbn = RTrim(ZTM.î≠íçïiî‘)
'            iB3_TKansn = ZTM.é¿ê—ä∑éZíl                                                ' 2009.02.27 del : é¿ê—ä∑éZílä÷òA
' ===================================================== *** 2008/06/27 End
'---------------------------------------------------------------------------------------' 2009.02.27 add start
                iB3_STorcd = ZTM.â¡çHêÊ
                Call iB3_STorcd_LostFocus
'                If ZTM.â¡çHêÊîwåiêF = 0 Then                       2009/12/08
'                   B3Btn_Hikei.BackColor = mBackColor
'                Else
'                   B3Btn_Hikei.BackColor = ZTM.â¡çHêÊîwåiêF
'                End If
                iB3_TTorcd = ZTM.ìæà”êÊ
                Call iB3_TTorcd_LostFocus
                iB3_SHinbn = ZTM.é©é–ïiî‘
                B3Chk_KanBn = ZTM.Ç©ÇÒÇŒÇÒÉTÉCÉY
                B3Chk_GenPn = ZTM.åªïiï[ÉTÉCÉY
                iB3_SHin1 = ZTM.égópïiî‘1
                iB3_SHin2 = ZTM.égópïiî‘2
                Call vsGridH_Select
'---------------------------------------------------------------------------------------' 2009.02.27 add end
    
        End If
    
        '============================================ ***** 2008/06/11 Start
        Call B3Comb_THKbn_Click
        '============================================ ***** 2008/06/11 End
    End If

End Sub
'+---------------------------------------+
'+  ÉOÉäÉbÉh (Click) éËîzêÊèÓïÒ
'+---------------------------------------+
Private Sub VSGridT_Click()
    With VsGridT
        If ProcHB = "B2" Then
           If Not Body1Chk Then Exit Sub
        End If
        
        If ProcHB = "B1" Or ProcHB = "B2" Then
            GridBack = ProcHB
            Call HeadBodyClear("B2")
            Call HeadBodyClear("B3")
            iB2_Torcd.Tag = .Row
            If Trim(.TextMatrix(.Row, 1)) <> "" Then
               Call GridToTextT
            Else                                ' 2008/02/06 ADD
               B2Op_Hinmoku(1).Value = True     ' 2008/02/06 ADD
            End If
        End If
        
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
        If ProcHB = "B2" Then
           B2_Area1.SetFocus                '>åüçıÇÃÃß›∏ºÆ›Ç™ñ≥ï\é¶Ç…Ç»ÇÈÇÃÇñhÇÆà◊
           iB2_Torcd.SetFocus
        End If
        If ProcHB = "B3" Then
           If SyoriKB2 = "ADD" Then
              Call GridToTextT
           End If
           iB3_Seban.SetFocus
        End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
    End With
    
'    Call vsGridH_Select                    ' 2009.03.03 del    ' 2009.03.02 add

End Sub
'+---------------------------------------+
'+  ÉOÉäÉbÉh (DblClick) éËîzêÊèÓïÒ
'+---------------------------------------+
Private Sub vsGridT_DblClick()
    
    With VsGridT
        If ProcHB = "B1" Or ProcHB = "B2" Then
                
            If Not Body1Chk Then Exit Sub
            
            If Trim(.TextMatrix(.Row, 1)) <> "" Then
                If SyoriKB2 <> "DEL" Then
                    RowColFlg = True
                    Call CisVsGridT.vsColor(å¬ï )
                    RowColFlg = False
                    
                    If CisFun.Val2(PB_Hin.Tag) <> 0 Then
                       Call PB_Hin_Click
                    End If
                    Call VSGridT_Click
                    SyoriKB2 = "MNT"
                    Call SyoriName2(SyoriKB2)
                    ProcHB = "B3"
                    Call DispChange(ProcHB)
                    ' ∏ﬁÿØƒﬁì‡óeï\é¶å„ÅA√∑Ωƒì¸óÕ
'                    Call CisVsGridS.EditVisible("C")
                    Call vsGridH_Select                 ' 2009.03.03 add
                    
                    iB3_Seban.SetFocus
                End If
            Else
                If ProcHB = "B1" Then
                   If CisFun.Val2(PB_Hin.Tag) <> 0 Then: Call PB_Hin_Click
                   Call PB_CHG_Click
                End If
            End If
        End If
    End With
End Sub

Private Sub VsGridT_GotFocus()
    If Not FormAct Then Exit Sub
    If RowColFlg Then Exit Sub
    
    Call VSGridT_Click
End Sub

'+---------------------------------------+
'+  ÉOÉäÉbÉh (GotFocus) éËîzêÊèÓïÒ
'+---------------------------------------+
Private Sub vsGridT_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
       Call VSGridT_Click
       Call vsGridT_DblClick
    End If
End Sub
'+---------------------------------------+
'+  ÉOÉäÉbÉh (RowColChange) éËîzêÊèÓïÒ
'+---------------------------------------+
Private Sub vsGridT_RowColChange()
    If Not FormAct Then Exit Sub
    If RowColFlg Then Exit Sub
    Call VSGridT_Click
End Sub

'****************************************************************************************'
'*      ÉOÉäÉbÉhì‡óeÉZÉbÉg(égópïiî‘èÓïÒ)
'****************************************************************************************'
Private Sub GridSetH()
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

' éËîzêÊÅ@ÉZÉbÉg
    gSL_Select = "SELECT ZH.*, HN.é–ì‡îwî‘ îwî‘çÜ,HN.ïiî‘ï“èW"
    gSL_Select = gSL_Select & " FROM çﬁóøïiî‘É}ÉXÉ^ ZH "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN ïiî‘É}ÉXÉ^ HN"
    gSL_Select = gSL_Select & "   ON HN.ïiî‘ = ZH.égópïiî‘ "
    gSL_Select = gSL_Select & " WHERE çﬁóøä«óùî‘çÜ = " & wKanri
    Call ZHMRead(gSL_Select, , 1)

' íäèoèàóù èâä˙æØƒ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' íäèoåèêî(çáåv)ëﬁî
' ∏ﬁÿØƒﬁçsçƒæØƒ & óÒà íuéÊìæ
    With VsGridH
       .Redraw = False
        If Ttl_Kensu + 1 < CisVsGridH.DispMax Then
            .Rows = CisVsGridH.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    
    Do Until Not ZHM_RDSTS
        ID = ID + 1
            
        Call HinInfoToGrid(ID)
        
        Call ZHMReadNext(1)
    Loop
' ÉNÉçÅ[ÉY
    Call ZHMClose

    With VsGridH
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'+----------------------------------------------------------------------------------+
'+      égópïiî‘èÓïÒÇÊÇËÉOÉäÉbÉhÇ÷ÉZÉbÉg
'+----------------------------------------------------------------------------------+
Private Sub HinInfoToGrid(SetRow As Long)
    With VsGridH
        For mIx = 0 To .Cols - 1
            .TextMatrix(SetRow, mIx) = ""
        Next mIx
        
        ' No
        .TextMatrix(SetRow, 0) = SetRow
        ' égópïiî‘
        .TextMatrix(SetRow, SColH(1)) = EditHinbn(ZHM.égópïiî‘, ZHM.ïiî‘ï“èW)
        ' é–ì‡îwî‘
        .TextMatrix(SetRow, SColH(2)) = RTrim(ZHM.îwî‘çÜ)
        ' ÉsÉbÉ`
        If ZHM.ÉsÉbÉ` <> 0 Then
            .TextMatrix(SetRow, SColH(3)) = CisFun.RSetFld(ZHM.ÉsÉbÉ`, 7, "###0.00")
        End If
        ' éÊêî
        If ZHM.éÊêî <> 0 Then
            .TextMatrix(SetRow, SColH(4)) = CisFun.RSetFld(ZHM.éÊêî, 4, "####")
        End If
        ' ñ¢égópó 
        If ZHM.ñ¢égópó  <> 0 Then
            .TextMatrix(SetRow, SColH(5)) = CisFun.RSetFld(ZHM.ñ¢égópó , 7, "###0.00")
        End If
        ' égópó 
        If ZHM.ñ¢égópó  <> 0 Then
            .TextMatrix(SetRow, SColH(6)) = CisFun.RSetFld(ZHM.égópó , 7, "###0.00")
        End If
        
        ' óLå¯ä˙å¿
        gStr = ""
        If Trim(ZHM.óLå¯ä˙å¿äJén) <> "" Then
            gStr = Mid(ZHM.óLå¯ä˙å¿äJén, 1, 4) & "/" & _
                   Mid(ZHM.óLå¯ä˙å¿äJén, 5, 2) & "/" & _
                   Mid(ZHM.óLå¯ä˙å¿äJén, 7, 2)
        End If
        If Trim(ZHM.óLå¯ä˙å¿èIóπ) <> "" Then
           If gStr = "" Then gStr = Space(10)
           gStr = gStr & "Å`"
           gStr = gStr & Mid(ZHM.óLå¯ä˙å¿èIóπ, 1, 4) & "/" & _
                         Mid(ZHM.óLå¯ä˙å¿èIóπ, 5, 2) & "/" & _
                         Mid(ZHM.óLå¯ä˙å¿èIóπ, 7, 2)
        Else
           If Trim(ZHM.óLå¯ä˙å¿äJén) <> "" Then
              gStr = gStr & "Å`" & Space(10)
           
           End If
        End If
        .TextMatrix(SetRow, SColH(7)) = gStr
'-------------------------------------------------------------------------------' 2009.03.03 add start
        ' égópïiî‘
        .TextMatrix(SetRow, SColH(9)) = ZHM.égópïiî‘
'-------------------------------------------------------------------------------' 2009.03.03 add end
    End With
End Sub
'-------------------------------------------------------------------------------' 2009.03.03 add start
'+---------------------------------------+
'+  ÉOÉäÉbÉh (Click) éËîzêÊèÓïÒ
'+---------------------------------------+
Private Sub VSGridH_Click()
    Dim Cnt_X As Integer
    Dim Cnt_L As Integer
'
'                RowColFlg = True
'                Call CisVsGridH.vsColor(å¬ï )
'                RowColFlg = False
'
'Exit Sub

    With VsGridH
        If ProcHB = "B3" Or ProcHB = "B1" Or ProcHB = "B2" Then
            If Trim(.TextMatrix(.Row, 1)) <> "" Then
                If Trim(.TextMatrix(.Row, 8)) <> "" Then
                    .TextMatrix(.Row, 8) = ""
'                    RowColFlg = True
'                    Call CisVsGridH.vsColor(è¡ãéå¬ï )
'                    RowColFlg = False
                Else
                    Cnt_X = 0
                    For Cnt_L = 1 To .Rows - 1
                        If Trim(.TextMatrix(Cnt_L, 8)) = "X" Then Cnt_X = Cnt_X + 1
                    Next Cnt_L
                    If Cnt_X >= 2 Then Exit Sub
                    .TextMatrix(.Row, 8) = "X"
                End If
                
                RowColFlg = True
                Call CisVsGridH.vsColor(å¬ï )
                RowColFlg = False
            End If
        End If
    End With
End Sub
'-------------------------------------------------------------------------------' 2009.03.03 add end
'-------------------------------------------------------------------------------' 2009.03.03 del start
''+---------------------------------------+
''+  ÉOÉäÉbÉh (Click) éËîzêÊèÓïÒ
''+---------------------------------------+
'Private Sub VSGridH_Click()
''    With VsGridH
''        If ProcHB = "B1" Or ProcHB = "B7" Then
''            GridBack = ProcHB
'''            Call HeadBodyClear("H2")
''            Call HeadBodyClear("B7")
''            Call HeadBodyClear("B8")
''            iB7_Hinbn.Tag = .Row
''            If Trim(.TextMatrix(.Row, 1)) <> "" Then
''               Call GridToTextH
''            End If
''        End If
''    End With
'End Sub
'-------------------------------------------------------------------------------' 2009.03.03 del end
'+---------------------------------------+
'+  ÉOÉäÉbÉh (DblClick) éËîzêÊèÓïÒ
'+---------------------------------------+
Private Sub vsGridH_DblClick()
   
'    With VsGridH
'        If ProcHB = "B1" Or ProcHB = "B7" Then
'
'            If Not Body1Chk Then Exit Sub
'
'            If Trim(.TextMatrix(.Row, 1)) <> "" Then
'                If SyoriKB3 <> "DEL" Then
'                    RowColFlg = True
'
'                    RowColFlg = False
'                    If CisFun.Val2(PB_Hin.Tag) = 0 Then: Call PB_Hin_Click
'                    Call VSGridH_Click
''                    Call CisVsGridH.vsColor(å¬ï )
'                    ProcHB = "B8"
'                    Call DispChange(ProcHB)
'                    SyoriKB3 = "MNT"
'                    Call SyoriName3(SyoriKB3)
'                End If
'            Else
'                If ProcHB = "B1" Then
'                   If CisFun.Val2(PB_Hin.Tag) = 0 Then: Call PB_Hin_Click
'                   Call PB_Chg_Click
'                End If
'            End If
'        End If
'    End With
End Sub
'+---------------------------------------+
'+  ÉOÉäÉbÉh (GotFocus) éËîzêÊèÓïÒ
'+---------------------------------------+
Private Sub vsGridH_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
       Call VSGridH_Click
       Call vsGridH_DblClick
    End If
End Sub
'+---------------------------------------+
'+  ÉOÉäÉbÉh (RowColChange) éËîzêÊèÓïÒ
'+---------------------------------------+
Private Sub vsGridH_RowColChange()
    If Not FormAct Then Exit Sub
    If RowColFlg Then Exit Sub
'    Call VSGridH_Click                         ' 2009.03.02 del
End Sub
'-----------------------------------------------' 2009.03.02 add start
'+---------------------------------------+
'+  égópïiî‘èÓïÒèâä˙ëIë
'+---------------------------------------+
Private Sub vsGridH_Select()
    Dim Cnt_L As Integer
    Dim Cnt_X As Integer
    
'    If ProcHB <> "B3" And ProcHB <> "B1" Then
'       Exit Sub
'    End If
    
    With VsGridH
        For gCnt1 = 1 To .Rows - 1
            .TextMatrix(gCnt1, 8) = ""
        Next
    End With
    RowColFlg = True
    Call CisVsGridH.vsColor(è¡ãéëSïî)
    RowColFlg = False
    
    Cnt_X = 0
    With VsGridH
        For Cnt_L = 1 To .Rows - 1
            If Trim(.TextMatrix(Cnt_L, 1)) = "" Then Exit For
            If (ZTM.égópïiî‘1 = Trim(.TextMatrix(Cnt_L, 9))) _
            Or (ZTM.égópïiî‘2 = Trim(.TextMatrix(Cnt_L, 9))) Then
              '.TextMatrix(Cnt_L, 8) = "X"
               .Row = Cnt_L
               Call VSGridH_Click
               Cnt_X = Cnt_X + 1
            Else
               '.TextMatrix(Cnt_L, 8) = ""
            End If
        Next Cnt_L
        
        If Cnt_X = 0 Then
           If .TextMatrix(1, 1) <> "" Then
             '.TextMatrix(1, 8) = "X"
              .Row = 1
              Call VSGridH_Click
           End If
           If .TextMatrix(2, 1) <> "" Then
             '.TextMatrix(2, 8) = "X"
              .Row = 2
              Call VSGridH_Click
           End If
        End If
    End With
End Sub
'-----------------------------------------------' 2009.03.02 add end

''+----------------------+
''+  VSGrid1_Scroll
''+----------------------+
'Private Sub vsGrid1_Scroll()
'    'ÉXÉNÉçÅ[ÉãéûÇÕÉeÉLÉXÉgÉ{ÉbÉNÉXÇîÒï\é¶Ç…ÇµÇ‹Ç∑ÅB
'    Call CisVsGrid.EditEnable
'End Sub
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
        .âpêîéö1 = "NJM "
        .êßå‰ãÊï™ = "X"
        .çÏê¨é“ = gTanto
        .åüçıñº = "éË îz ãÊ ï™ åü çı"
        .ÉRÅ[Éhñº = "éËîzãÊ"
        .ñºèÃñº = "éË îz ãÊ ï™ ñº"
        If Not SYMInitCreate(True) Then
            .íl = "1"
            .ílñºèÃ = "íºëó"
            .âpêîéö1 = "NJMS"
            Call SYMInitCreate
            .íl = "2"
            .ílñºèÃ = "é–ì‡"
            .âpêîéö1 = "  M "
            Call SYMInitCreate
            .íl = "3"
            .ílñºèÃ = "éÛãã"
            .âpêîéö1 = "NJM "
            Call SYMInitCreate
            '=================================***** 2008/06/11 Start
            .íl = "4"
            .ílñºèÃ = "éxãã"
            .âpêîéö1 = "N"
            Call SYMInitCreate
            '=================================***** 2008/06/11 End
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
            .ílñºèÃ = "åWêî"
            Call SYMInitCreate
            .íl = "2"
            .ílñºèÃ = "î≠íçì_"
            Call SYMInitCreate
            .íl = "3"
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
' çﬁéÌãÊï™
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "çﬁéÌãÊï™"
        .ãÊï™åÖêî = 2
        .íl = "10"
        .ílñºèÃ = "ÉXÉeÉìÉåÉX"
        .êßå‰ãÊï™ = ""
        .çÏê¨é“ = gTanto
        .åüçıñº = "çﬁ éÌ ãÊ ï™ åü çı"
        .ÉRÅ[Éhñº = "çﬁéÌãÊ"
        .ñºèÃñº = "çﬁ éÌ ãÊ ï™ ñº"
        If Not SYMInitCreate(True) Then
            .íl = "99"
            .ílñºèÃ = "ÇªÇÃëº"
            Call SYMInitCreate
        End If
    End With
' å`èÛãÊï™
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "å`èÛãÊï™"
        .ãÊï™åÖêî = 2
        .íl = "10"
        .ílñºèÃ = "ÉRÉCÉã"
        .êßå‰ãÊï™ = ""
        .çÏê¨é“ = gTanto
        .åüçıñº = "å` èÛ ãÊ ï™ åü çı"
        .ÉRÅ[Éhñº = "å`èÛãÊ"
        .ñºèÃñº = "å` èÛ ãÊ ï™ ñº"
        .âpêîéö3 = "C"
        If Not SYMInitCreate(True) Then
            .íl = "99"
            .ílñºèÃ = "ÇªÇÃëº"
            Call SYMInitCreate
        End If
    End With
' íPà ãÊï™
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "íPà ãÊï™"
        .ãÊï™åÖêî = 2
        .íl = "10"
        .ílñºèÃ = "å¬"
        .êßå‰ãÊï™ = ""
        .çÏê¨é“ = gTanto
        .åüçıñº = "íP à  ãÊ ï™ åü çı"
        .ÉRÅ[Éhñº = "íPà ãÊ"
        .ñºèÃñº = "íP à  ãÊ ï™ ñº"
        If Not SYMInitCreate(True) Then
            .íl = "11"
            .ílñºèÃ = "ä™"
            Call SYMInitCreate
            .íl = "20"
            .ílñºèÃ = "ár"
            Call SYMInitCreate
            .íl = "21"
            .ílñºèÃ = "Çá"
            Call SYMInitCreate
            .íl = "22"
            .ílñºèÃ = "ás"
            Call SYMInitCreate
            .íl = "23"
            .ílñºèÃ = "Çî"
            Call SYMInitCreate
            .íl = "25"
            .ílñºèÃ = "áo"
            Call SYMInitCreate
            .íl = "26"
            .ílñºèÃ = "áp"
            Call SYMInitCreate
            .íl = "27"
            .ílñºèÃ = "Çç"
            Call SYMInitCreate
            .íl = "28"
            .ílñºèÃ = "áq"
            Call SYMInitCreate
        End If
    End With
End Sub


