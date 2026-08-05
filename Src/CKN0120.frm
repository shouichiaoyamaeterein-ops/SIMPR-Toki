VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Object = "{543CD9A5-1043-4C19-96AC-FC4793F212CA}#6.0#0"; "CISYMD~2.OCX"
Begin VB.Form CKN0120 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "î≠íçÉeÅ[ÉuÉãÅ@ÉÅÉìÉe"
   ClientHeight    =   11295
   ClientLeft      =   2550
   ClientTop       =   3435
   ClientWidth     =   15540
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
   ScaleHeight     =   11295
   ScaleWidth      =   15540
   WindowState     =   2  'ç≈ëÂâª
   Begin Cis3D_v60.CIS3D PNL_Wait 
      Height          =   930
      Left            =   4920
      Top             =   5280
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
         Height          =   675
         Left            =   150
         Top             =   135
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   1191
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "ÅyçXêVíÜÅzÇµÇŒÇÁÇ≠Ç®ë“Çøâ∫Ç≥Ç¢"
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Ç»Çµ
      Height          =   345
      Index           =   1
      Left            =   8025
      ScaleHeight     =   345
      ScaleWidth      =   1860
      TabIndex        =   154
      TabStop         =   0   'False
      Top             =   885
      Width           =   1860
      Begin Cis3D_v60.CIS3D UC_3D6 
         Height          =   345
         Left            =   45
         Top             =   0
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   609
         BackColor       =   14737632
         ForeColor       =   0
         Caption         =   "ï\é¶åèêî"
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
         Begin Cis3D_v60.CIS3D iT1_Kensu 
            Height          =   285
            Left            =   1050
            Top             =   30
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   503
            BackColor       =   16777152
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
   Begin VB.PictureBox H1_Area1 
      Height          =   2070
      Left            =   75
      ScaleHeight     =   2010
      ScaleWidth      =   15165
      TabIndex        =   153
      TabStop         =   0   'False
      Top             =   1275
      Width           =   15225
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1935
         Left            =   45
         Top             =   60
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   3413
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
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   1755
            Left            =   105
            Top             =   90
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   3096
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
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   855
               Left            =   15
               Top             =   885
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   1508
               Caption         =   "ç≈èIî≠íçèàóùì˙"
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
               Begin Cis3D_v60.CIS3D CIS3D8 
                  Height          =   315
                  Left            =   2025
                  Top             =   465
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   556
                  Caption         =   "âÒñ⁄"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9
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
               Begin Cis3D_v60.CIS3D H1lb_EHYmd 
                  Height          =   360
                  Left            =   105
                  Top             =   345
                  Width           =   1380
                  _ExtentX        =   2434
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
               Begin Cis3D_v60.CIS3D H1lb_EHKai 
                  Height          =   360
                  Left            =   1500
                  Top             =   345
                  Width           =   480
                  _ExtentX        =   847
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cAlingnment     =   7
                  cBoderStyle     =   1
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   870
               Left            =   15
               Top             =   15
               Width           =   3570
               _ExtentX        =   6297
               _ExtentY        =   1535
               Caption         =   "îıÅ@çl"
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
               Begin CisText_V60.CisText iH1_Bikou 
                  Height          =   360
                  Left            =   105
                  TabIndex        =   112
                  Top             =   375
                  Width           =   3390
                  _ExtentX        =   5980
                  _ExtentY        =   635
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   40
                  IMEMode         =   4
                  cDataType       =   1
                  cDataReplace    =   1
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D13 
               Height          =   855
               Left            =   2460
               Top             =   885
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   1508
               Caption         =   "î≠íçâÒêî"
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
               Begin Cis3D_v60.CIS3D CIS3D14 
                  Height          =   315
                  Left            =   1305
                  Top             =   450
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   556
                  Caption         =   "âÒ"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9
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
               Begin Cis3D_v60.CIS3D H1lb_MaxKai 
                  Height          =   360
                  Left            =   300
                  Top             =   345
                  Width           =   480
                  _ExtentX        =   847
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cAlingnment     =   7
                  cBoderStyle     =   1
               End
            End
         End
         Begin Cis3D_v60.CIS3D H1_Area 
            Height          =   1785
            Index           =   4
            Left            =   3735
            Top             =   75
            Width           =   11070
            _ExtentX        =   19526
            _ExtentY        =   3149
            BackColor       =   0
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   1
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   1
               Left            =   1185
               Top             =   465
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_KHh 
                  Height          =   345
                  Index           =   1
                  Left            =   15
                  TabIndex        =   32
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_KMm 
                  Height          =   345
                  Index           =   1
                  Left            =   300
                  TabIndex        =   33
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D3 
               Height          =   420
               Index           =   4
               Left            =   30
               Top             =   465
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   741
               ForeColor       =   16711680
               Caption         =   "âÒé˚éûä‘"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   32
               Left            =   1185
               Top             =   30
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   16761024
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_HHh 
                  Height          =   345
                  Index           =   1
                  Left            =   15
                  TabIndex        =   0
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_HMm 
                  Height          =   345
                  Index           =   1
                  Left            =   300
                  TabIndex        =   1
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D3 
               Height          =   420
               Index           =   5
               Left            =   30
               Top             =   30
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   741
               BackColor       =   16761024
               ForeColor       =   192
               Caption         =   "î≠íçéûä‘"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   40
               Left            =   1185
               Top             =   900
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_THh 
                  Height          =   345
                  Index           =   1
                  Left            =   15
                  TabIndex        =   64
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_TMm 
                  Height          =   345
                  Index           =   1
                  Left            =   300
                  TabIndex        =   65
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D3 
               Height          =   420
               Index           =   6
               Left            =   30
               Top             =   900
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   741
               ForeColor       =   16711680
               Caption         =   "à¯éÊéûä‘"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   420
               Index           =   8
               Left            =   1185
               Top             =   1335
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               Begin VB.CheckBox ChkH1_HYKb 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "íã"
                  BeginProperty Font 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   60
                  TabIndex        =   96
                  Top             =   90
                  Width           =   525
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D3 
               Height          =   420
               Index           =   7
               Left            =   30
               Top             =   1335
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   741
               ForeColor       =   16711680
               Caption         =   "íãñÈãŒãÊ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   4
               Left            =   1800
               Top             =   465
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_KHh 
                  Height          =   345
                  Index           =   2
                  Left            =   15
                  TabIndex        =   34
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_KMm 
                  Height          =   345
                  Index           =   2
                  Left            =   300
                  TabIndex        =   35
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   5
               Left            =   1800
               Top             =   30
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   16761024
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_HHh 
                  Height          =   345
                  Index           =   2
                  Left            =   15
                  TabIndex        =   2
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_HMm 
                  Height          =   345
                  Index           =   2
                  Left            =   300
                  TabIndex        =   3
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   6
               Left            =   1800
               Top             =   900
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_THh 
                  Height          =   345
                  Index           =   2
                  Left            =   15
                  TabIndex        =   66
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_TMm 
                  Height          =   345
                  Index           =   2
                  Left            =   300
                  TabIndex        =   67
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   420
               Index           =   1
               Left            =   1800
               Top             =   1335
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               Begin VB.CheckBox ChkH1_HYKb 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "íã"
                  BeginProperty Font 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   60
                  TabIndex        =   97
                  Top             =   90
                  Width           =   525
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   7
               Left            =   2415
               Top             =   465
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_KHh 
                  Height          =   345
                  Index           =   3
                  Left            =   15
                  TabIndex        =   36
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_KMm 
                  Height          =   345
                  Index           =   3
                  Left            =   300
                  TabIndex        =   37
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   8
               Left            =   2415
               Top             =   30
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   16761024
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_HHh 
                  Height          =   345
                  Index           =   3
                  Left            =   15
                  TabIndex        =   4
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_HMm 
                  Height          =   345
                  Index           =   3
                  Left            =   300
                  TabIndex        =   5
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   9
               Left            =   2415
               Top             =   900
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_THh 
                  Height          =   345
                  Index           =   3
                  Left            =   15
                  TabIndex        =   68
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_TMm 
                  Height          =   345
                  Index           =   3
                  Left            =   300
                  TabIndex        =   69
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   420
               Index           =   2
               Left            =   2415
               Top             =   1335
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               Begin VB.CheckBox ChkH1_HYKb 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "íã"
                  BeginProperty Font 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   3
                  Left            =   60
                  TabIndex        =   98
                  Top             =   90
                  Width           =   525
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   10
               Left            =   3030
               Top             =   465
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_KHh 
                  Height          =   345
                  Index           =   4
                  Left            =   15
                  TabIndex        =   38
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_KMm 
                  Height          =   345
                  Index           =   4
                  Left            =   300
                  TabIndex        =   39
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   11
               Left            =   3030
               Top             =   30
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   16761024
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_HHh 
                  Height          =   345
                  Index           =   4
                  Left            =   15
                  TabIndex        =   6
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_HMm 
                  Height          =   345
                  Index           =   4
                  Left            =   300
                  TabIndex        =   7
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   12
               Left            =   3030
               Top             =   900
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_THh 
                  Height          =   345
                  Index           =   4
                  Left            =   15
                  TabIndex        =   70
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_TMm 
                  Height          =   345
                  Index           =   4
                  Left            =   300
                  TabIndex        =   71
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   420
               Index           =   3
               Left            =   3030
               Top             =   1335
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               Begin VB.CheckBox ChkH1_HYKb 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "íã"
                  BeginProperty Font 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   4
                  Left            =   60
                  TabIndex        =   99
                  Top             =   90
                  Width           =   525
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   13
               Left            =   3645
               Top             =   465
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_KHh 
                  Height          =   345
                  Index           =   5
                  Left            =   15
                  TabIndex        =   40
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_KMm 
                  Height          =   345
                  Index           =   5
                  Left            =   300
                  TabIndex        =   41
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   14
               Left            =   3645
               Top             =   30
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   16761024
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_HHh 
                  Height          =   345
                  Index           =   5
                  Left            =   15
                  TabIndex        =   8
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_HMm 
                  Height          =   345
                  Index           =   5
                  Left            =   300
                  TabIndex        =   9
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   15
               Left            =   3645
               Top             =   900
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_THh 
                  Height          =   345
                  Index           =   5
                  Left            =   15
                  TabIndex        =   72
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_TMm 
                  Height          =   345
                  Index           =   5
                  Left            =   300
                  TabIndex        =   73
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   420
               Index           =   4
               Left            =   3645
               Top             =   1335
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               Begin VB.CheckBox ChkH1_HYKb 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "íã"
                  BeginProperty Font 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   5
                  Left            =   60
                  TabIndex        =   100
                  Top             =   90
                  Width           =   525
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   16
               Left            =   4260
               Top             =   465
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_KHh 
                  Height          =   345
                  Index           =   6
                  Left            =   15
                  TabIndex        =   42
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_KMm 
                  Height          =   345
                  Index           =   6
                  Left            =   300
                  TabIndex        =   43
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   17
               Left            =   4260
               Top             =   30
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   16761024
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_HHh 
                  Height          =   345
                  Index           =   6
                  Left            =   15
                  TabIndex        =   10
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_HMm 
                  Height          =   345
                  Index           =   6
                  Left            =   300
                  TabIndex        =   11
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   18
               Left            =   4260
               Top             =   900
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_THh 
                  Height          =   345
                  Index           =   6
                  Left            =   15
                  TabIndex        =   74
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_TMm 
                  Height          =   345
                  Index           =   6
                  Left            =   300
                  TabIndex        =   75
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   420
               Index           =   5
               Left            =   4260
               Top             =   1335
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               Begin VB.CheckBox ChkH1_HYKb 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "íã"
                  BeginProperty Font 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   6
                  Left            =   60
                  TabIndex        =   101
                  Top             =   90
                  Width           =   525
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   19
               Left            =   4875
               Top             =   465
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_KHh 
                  Height          =   345
                  Index           =   7
                  Left            =   15
                  TabIndex        =   44
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_KMm 
                  Height          =   345
                  Index           =   7
                  Left            =   300
                  TabIndex        =   45
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   20
               Left            =   4875
               Top             =   30
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   16761024
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_HHh 
                  Height          =   345
                  Index           =   7
                  Left            =   15
                  TabIndex        =   12
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_HMm 
                  Height          =   345
                  Index           =   7
                  Left            =   300
                  TabIndex        =   13
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   21
               Left            =   4875
               Top             =   900
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_THh 
                  Height          =   345
                  Index           =   7
                  Left            =   15
                  TabIndex        =   76
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_TMm 
                  Height          =   345
                  Index           =   7
                  Left            =   300
                  TabIndex        =   77
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   420
               Index           =   6
               Left            =   4875
               Top             =   1335
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               Begin VB.CheckBox ChkH1_HYKb 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "íã"
                  BeginProperty Font 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   7
                  Left            =   60
                  TabIndex        =   102
                  Top             =   90
                  Width           =   525
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   22
               Left            =   5490
               Top             =   465
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_KHh 
                  Height          =   345
                  Index           =   8
                  Left            =   15
                  TabIndex        =   46
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_KMm 
                  Height          =   345
                  Index           =   8
                  Left            =   300
                  TabIndex        =   47
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   23
               Left            =   5490
               Top             =   30
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   16761024
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_HHh 
                  Height          =   345
                  Index           =   8
                  Left            =   15
                  TabIndex        =   14
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_HMm 
                  Height          =   345
                  Index           =   8
                  Left            =   300
                  TabIndex        =   15
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   24
               Left            =   5490
               Top             =   900
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_THh 
                  Height          =   345
                  Index           =   8
                  Left            =   15
                  TabIndex        =   78
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_TMm 
                  Height          =   345
                  Index           =   8
                  Left            =   300
                  TabIndex        =   79
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   420
               Index           =   7
               Left            =   5490
               Top             =   1335
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               Begin VB.CheckBox ChkH1_HYKb 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "íã"
                  BeginProperty Font 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   8
                  Left            =   60
                  TabIndex        =   103
                  Top             =   90
                  Width           =   525
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   25
               Left            =   6105
               Top             =   465
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_KHh 
                  Height          =   345
                  Index           =   9
                  Left            =   15
                  TabIndex        =   48
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_KMm 
                  Height          =   345
                  Index           =   9
                  Left            =   300
                  TabIndex        =   49
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   26
               Left            =   6105
               Top             =   30
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   16761024
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_HHh 
                  Height          =   345
                  Index           =   9
                  Left            =   15
                  TabIndex        =   16
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_HMm 
                  Height          =   345
                  Index           =   9
                  Left            =   300
                  TabIndex        =   17
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   27
               Left            =   6105
               Top             =   900
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_THh 
                  Height          =   345
                  Index           =   9
                  Left            =   15
                  TabIndex        =   80
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_TMm 
                  Height          =   345
                  Index           =   9
                  Left            =   300
                  TabIndex        =   81
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   420
               Index           =   9
               Left            =   6105
               Top             =   1335
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               Begin VB.CheckBox ChkH1_HYKb 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "íã"
                  BeginProperty Font 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   9
                  Left            =   60
                  TabIndex        =   104
                  Top             =   90
                  Width           =   525
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   28
               Left            =   6720
               Top             =   465
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_KHh 
                  Height          =   345
                  Index           =   10
                  Left            =   15
                  TabIndex        =   50
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_KMm 
                  Height          =   345
                  Index           =   10
                  Left            =   300
                  TabIndex        =   51
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   29
               Left            =   6720
               Top             =   30
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   16761024
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_HHh 
                  Height          =   345
                  Index           =   10
                  Left            =   15
                  TabIndex        =   18
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_HMm 
                  Height          =   345
                  Index           =   10
                  Left            =   300
                  TabIndex        =   19
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   30
               Left            =   6720
               Top             =   900
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_THh 
                  Height          =   345
                  Index           =   10
                  Left            =   15
                  TabIndex        =   82
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_TMm 
                  Height          =   345
                  Index           =   10
                  Left            =   300
                  TabIndex        =   83
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   420
               Index           =   10
               Left            =   6720
               Top             =   1335
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               Begin VB.CheckBox ChkH1_HYKb 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "íã"
                  BeginProperty Font 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   10
                  Left            =   60
                  TabIndex        =   105
                  Top             =   90
                  Width           =   525
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   31
               Left            =   7335
               Top             =   465
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_KHh 
                  Height          =   345
                  Index           =   11
                  Left            =   15
                  TabIndex        =   52
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_KMm 
                  Height          =   345
                  Index           =   11
                  Left            =   300
                  TabIndex        =   53
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   33
               Left            =   7335
               Top             =   30
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   16761024
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_HHh 
                  Height          =   345
                  Index           =   11
                  Left            =   15
                  TabIndex        =   20
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_HMm 
                  Height          =   345
                  Index           =   11
                  Left            =   300
                  TabIndex        =   21
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   34
               Left            =   7335
               Top             =   900
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_THh 
                  Height          =   345
                  Index           =   11
                  Left            =   15
                  TabIndex        =   84
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_TMm 
                  Height          =   345
                  Index           =   11
                  Left            =   300
                  TabIndex        =   85
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   420
               Index           =   11
               Left            =   7335
               Top             =   1335
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               Begin VB.CheckBox ChkH1_HYKb 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "íã"
                  BeginProperty Font 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   11
                  Left            =   60
                  TabIndex        =   106
                  Top             =   90
                  Width           =   525
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   35
               Left            =   7950
               Top             =   465
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_KHh 
                  Height          =   345
                  Index           =   12
                  Left            =   15
                  TabIndex        =   54
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_KMm 
                  Height          =   345
                  Index           =   12
                  Left            =   300
                  TabIndex        =   55
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   36
               Left            =   7950
               Top             =   30
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   16761024
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_HHh 
                  Height          =   345
                  Index           =   12
                  Left            =   15
                  TabIndex        =   22
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_HMm 
                  Height          =   345
                  Index           =   12
                  Left            =   300
                  TabIndex        =   23
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   37
               Left            =   7950
               Top             =   900
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_THh 
                  Height          =   345
                  Index           =   12
                  Left            =   15
                  TabIndex        =   86
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_TMm 
                  Height          =   345
                  Index           =   12
                  Left            =   300
                  TabIndex        =   87
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   420
               Index           =   12
               Left            =   7950
               Top             =   1335
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               Begin VB.CheckBox ChkH1_HYKb 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "íã"
                  BeginProperty Font 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   12
                  Left            =   60
                  TabIndex        =   107
                  Top             =   90
                  Width           =   525
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   38
               Left            =   8565
               Top             =   465
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_KHh 
                  Height          =   345
                  Index           =   13
                  Left            =   15
                  TabIndex        =   56
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_KMm 
                  Height          =   345
                  Index           =   13
                  Left            =   300
                  TabIndex        =   57
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   39
               Left            =   8565
               Top             =   30
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   16761024
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_HHh 
                  Height          =   345
                  Index           =   13
                  Left            =   15
                  TabIndex        =   24
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_HMm 
                  Height          =   345
                  Index           =   13
                  Left            =   300
                  TabIndex        =   25
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   41
               Left            =   8565
               Top             =   900
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_THh 
                  Height          =   345
                  Index           =   13
                  Left            =   15
                  TabIndex        =   88
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_TMm 
                  Height          =   345
                  Index           =   13
                  Left            =   300
                  TabIndex        =   89
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   420
               Index           =   13
               Left            =   8565
               Top             =   1335
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               Begin VB.CheckBox ChkH1_HYKb 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "íã"
                  BeginProperty Font 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   13
                  Left            =   60
                  TabIndex        =   108
                  Top             =   90
                  Width           =   525
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   42
               Left            =   9180
               Top             =   465
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_KHh 
                  Height          =   345
                  Index           =   14
                  Left            =   15
                  TabIndex        =   58
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_KMm 
                  Height          =   345
                  Index           =   14
                  Left            =   300
                  TabIndex        =   59
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   43
               Left            =   9180
               Top             =   30
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   16761024
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_HHh 
                  Height          =   345
                  Index           =   14
                  Left            =   15
                  TabIndex        =   26
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_HMm 
                  Height          =   345
                  Index           =   14
                  Left            =   300
                  TabIndex        =   27
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   44
               Left            =   9180
               Top             =   900
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_THh 
                  Height          =   345
                  Index           =   14
                  Left            =   15
                  TabIndex        =   90
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_TMm 
                  Height          =   345
                  Index           =   14
                  Left            =   300
                  TabIndex        =   91
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   420
               Index           =   14
               Left            =   9180
               Top             =   1335
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               Begin VB.CheckBox ChkH1_HYKb 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "íã"
                  BeginProperty Font 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   14
                  Left            =   60
                  TabIndex        =   109
                  Top             =   90
                  Width           =   525
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   45
               Left            =   9795
               Top             =   465
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_KHh 
                  Height          =   345
                  Index           =   15
                  Left            =   15
                  TabIndex        =   60
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_KMm 
                  Height          =   345
                  Index           =   15
                  Left            =   300
                  TabIndex        =   61
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   46
               Left            =   9795
               Top             =   30
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   16761024
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_HHh 
                  Height          =   345
                  Index           =   15
                  Left            =   15
                  TabIndex        =   28
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_HMm 
                  Height          =   345
                  Index           =   15
                  Left            =   300
                  TabIndex        =   29
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   47
               Left            =   9795
               Top             =   900
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_THh 
                  Height          =   345
                  Index           =   15
                  Left            =   15
                  TabIndex        =   92
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_TMm 
                  Height          =   345
                  Index           =   15
                  Left            =   300
                  TabIndex        =   93
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   420
               Index           =   15
               Left            =   9795
               Top             =   1335
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               Begin VB.CheckBox ChkH1_HYKb 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "íã"
                  BeginProperty Font 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   15
                  Left            =   60
                  TabIndex        =   110
                  Top             =   90
                  Width           =   525
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   48
               Left            =   10410
               Top             =   465
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_KHh 
                  Height          =   345
                  Index           =   16
                  Left            =   15
                  TabIndex        =   62
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_KMm 
                  Height          =   345
                  Index           =   16
                  Left            =   300
                  TabIndex        =   63
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   49
               Left            =   10410
               Top             =   30
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   16761024
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_HHh 
                  Height          =   345
                  Index           =   16
                  Left            =   15
                  TabIndex        =   30
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_HMm 
                  Height          =   345
                  Index           =   16
                  Left            =   300
                  TabIndex        =   31
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   50
               Left            =   10410
               Top             =   900
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   ":"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9
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
               Begin CisText_V60.CisText iH1_THh 
                  Height          =   345
                  Index           =   16
                  Left            =   15
                  TabIndex        =   94
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iH1_TMm 
                  Height          =   345
                  Index           =   16
                  Left            =   300
                  TabIndex        =   95
                  Top             =   45
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
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
                  cGFormat        =   "00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   420
               Index           =   16
               Left            =   10410
               Top             =   1335
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               Begin VB.CheckBox ChkH1_HYKb 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "íã"
                  BeginProperty Font 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   16
                  Left            =   60
                  TabIndex        =   111
                  Top             =   90
                  Width           =   525
               End
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
      Height          =   705
      Left            =   10140
      ScaleHeight     =   645
      ScaleWidth      =   5085
      TabIndex        =   151
      TabStop         =   0   'False
      Top             =   510
      Width           =   5145
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   2535
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
         Left            =   3375
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
         Left            =   4215
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
      Begin CisBtn_60.CisBtn PB_PRT 
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
      Begin CisBtn_60.CisBtn PB_MNT 
         Height          =   585
         Left            =   855
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
         BtnCaption      =   "í èÌ"
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
      Begin CisBtn_60.CisBtn PB_KAN 
         Height          =   585
         Left            =   1695
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
         BtnCaption      =   "ä»à’"
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
         PFBackColor     =   16711935
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
      TabIndex        =   150
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
      Height          =   7785
      Left            =   60
      ScaleHeight     =   7725
      ScaleWidth      =   15180
      TabIndex        =   131
      TabStop         =   0   'False
      Top             =   3375
      Width           =   15240
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   7620
         Index           =   1
         Left            =   30
         Top             =   60
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   13441
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
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin vsFlexLib.vsFlexArray vsGrid 
            Height          =   7470
            Left            =   45
            TabIndex        =   152
            TabStop         =   0   'False
            Top             =   60
            Width           =   15015
            _Version        =   196608
            _ExtentX        =   26485
            _ExtentY        =   13176
            _StockProps     =   228
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "áÇ |édì¸êÊ |éÛì¸|ÉTÉCÉNÉã|ç≈èIî≠íçì˙-ï÷   |xxâÒ  |xxâÒ  |xxâÒ  |xxâÒ  |xxâÒ  |xxâÒ  |xxâÒ  |xxâÒ  ||"
            Rows            =   17
            Cols            =   63
            BackColor       =   12648384
            BackColorBkg    =   12632256
            FocusRect       =   2
            HighLight       =   2
            RowHeightMin    =   430
            SelectionMode   =   1
            Begin Cis3D_v60.CIS3D Input_Row 
               Height          =   435
               Left            =   3075
               Top             =   435
               Width           =   11850
               _ExtentX        =   20902
               _ExtentY        =   767
               BackColor       =   0
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cBoderStyle     =   1
               Begin Cis3D_v60.CIS3D CIS3D6 
                  Height          =   420
                  Left            =   0
                  Top             =   0
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   741
                  BackColor       =   14737632
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
                  Begin CisYMD2_V60.CisYMD2 iB1_EHYmd 
                     Height          =   345
                     Left            =   15
                     TabIndex        =   113
                     Top             =   30
                     Width           =   1410
                     _ExtentX        =   2487
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
                     Object.Height          =   345
                     cSize           =   -1  'True
                     cChkResult      =   0   'False
                  End
                  Begin CisText_V60.CisText iB1_EHBin 
                     Height          =   360
                     Left            =   1410
                     TabIndex        =   114
                     Top             =   30
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   635
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
                     cFdAutoFormat   =   1
                     cGFormat        =   "##"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D Back_Bin 
                  Height          =   420
                  Index           =   1
                  Left            =   1785
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   8
                  Begin VB.CheckBox ChkB1_ZHKbn 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Check1"
                     Height          =   345
                     Index           =   1
                     Left            =   390
                     TabIndex        =   116
                     Top             =   30
                     Visible         =   0   'False
                     Width           =   180
                  End
                  Begin CisText_V60.CisText iB1_Bin 
                     Height          =   360
                     Index           =   1
                     Left            =   30
                     TabIndex        =   115
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   635
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
                     cFdAutoFormat   =   1
                     cGFormat        =   "##"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D Back_Bin 
                  Height          =   420
                  Index           =   2
                  Left            =   2400
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   8
                  Begin VB.CheckBox ChkB1_ZHKbn 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Check1"
                     Height          =   345
                     Index           =   2
                     Left            =   390
                     TabIndex        =   118
                     Top             =   30
                     Visible         =   0   'False
                     Width           =   210
                  End
                  Begin CisText_V60.CisText iB1_Bin 
                     Height          =   360
                     Index           =   2
                     Left            =   30
                     TabIndex        =   117
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   635
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
                     cGFormat        =   "00"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D Back_Bin 
                  Height          =   420
                  Index           =   3
                  Left            =   3015
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   8
                  Begin VB.CheckBox ChkB1_ZHKbn 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Check1"
                     Height          =   345
                     Index           =   3
                     Left            =   390
                     TabIndex        =   120
                     Top             =   30
                     Visible         =   0   'False
                     Width           =   210
                  End
                  Begin CisText_V60.CisText iB1_Bin 
                     Height          =   360
                     Index           =   3
                     Left            =   30
                     TabIndex        =   119
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   635
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
                     cGFormat        =   "00"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D Back_Bin 
                  Height          =   420
                  Index           =   4
                  Left            =   3630
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   8
                  Begin VB.CheckBox ChkB1_ZHKbn 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Check1"
                     Height          =   345
                     Index           =   4
                     Left            =   390
                     TabIndex        =   122
                     Top             =   30
                     Visible         =   0   'False
                     Width           =   180
                  End
                  Begin CisText_V60.CisText iB1_Bin 
                     Height          =   360
                     Index           =   4
                     Left            =   30
                     TabIndex        =   121
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   635
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
                     cFdAutoFormat   =   1
                     cGFormat        =   "##"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D Back_Bin 
                  Height          =   420
                  Index           =   5
                  Left            =   4245
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   8
                  Begin VB.CheckBox ChkB1_ZHKbn 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Check1"
                     Height          =   345
                     Index           =   5
                     Left            =   390
                     TabIndex        =   124
                     Top             =   30
                     Visible         =   0   'False
                     Width           =   210
                  End
                  Begin CisText_V60.CisText iB1_Bin 
                     Height          =   360
                     Index           =   5
                     Left            =   30
                     TabIndex        =   123
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   635
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
                     cGFormat        =   "00"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D Back_Bin 
                  Height          =   420
                  Index           =   6
                  Left            =   4860
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   8
                  Begin VB.CheckBox ChkB1_ZHKbn 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Check1"
                     Height          =   345
                     Index           =   6
                     Left            =   390
                     TabIndex        =   126
                     Top             =   30
                     Visible         =   0   'False
                     Width           =   180
                  End
                  Begin CisText_V60.CisText iB1_Bin 
                     Height          =   360
                     Index           =   6
                     Left            =   30
                     TabIndex        =   125
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   635
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
                     cFdAutoFormat   =   1
                     cGFormat        =   "##"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D Back_Bin 
                  Height          =   420
                  Index           =   7
                  Left            =   5475
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   8
                  Begin VB.CheckBox ChkB1_ZHKbn 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Check1"
                     Height          =   345
                     Index           =   7
                     Left            =   390
                     TabIndex        =   128
                     Top             =   30
                     Visible         =   0   'False
                     Width           =   210
                  End
                  Begin CisText_V60.CisText iB1_Bin 
                     Height          =   360
                     Index           =   7
                     Left            =   30
                     TabIndex        =   127
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   635
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
                     cGFormat        =   "00"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D Back_Bin 
                  Height          =   420
                  Index           =   8
                  Left            =   6090
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   8
                  Begin VB.CheckBox ChkB1_ZHKbn 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Check1"
                     Height          =   345
                     Index           =   8
                     Left            =   390
                     TabIndex        =   130
                     Top             =   30
                     Visible         =   0   'False
                     Width           =   210
                  End
                  Begin CisText_V60.CisText iB1_Bin 
                     Height          =   360
                     Index           =   8
                     Left            =   30
                     TabIndex        =   129
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   635
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
                     cGFormat        =   "00"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D Back_Bin 
                  Height          =   420
                  Index           =   9
                  Left            =   6705
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   8
                  Begin VB.CheckBox ChkB1_ZHKbn 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Check1"
                     Height          =   345
                     Index           =   9
                     Left            =   390
                     TabIndex        =   133
                     Top             =   30
                     Visible         =   0   'False
                     Width           =   180
                  End
                  Begin CisText_V60.CisText iB1_Bin 
                     Height          =   360
                     Index           =   9
                     Left            =   30
                     TabIndex        =   132
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   635
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
                     cFdAutoFormat   =   1
                     cGFormat        =   "##"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D Back_Bin 
                  Height          =   420
                  Index           =   10
                  Left            =   7320
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   8
                  Begin VB.CheckBox ChkB1_ZHKbn 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Check1"
                     Height          =   345
                     Index           =   10
                     Left            =   390
                     TabIndex        =   135
                     Top             =   30
                     Visible         =   0   'False
                     Width           =   210
                  End
                  Begin CisText_V60.CisText iB1_Bin 
                     Height          =   360
                     Index           =   10
                     Left            =   30
                     TabIndex        =   134
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   635
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
                     cGFormat        =   "00"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D Back_Bin 
                  Height          =   420
                  Index           =   11
                  Left            =   7935
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   8
                  Begin VB.CheckBox ChkB1_ZHKbn 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Check1"
                     Height          =   345
                     Index           =   11
                     Left            =   390
                     TabIndex        =   137
                     Top             =   30
                     Visible         =   0   'False
                     Width           =   210
                  End
                  Begin CisText_V60.CisText iB1_Bin 
                     Height          =   360
                     Index           =   11
                     Left            =   30
                     TabIndex        =   136
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   635
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
                     cGFormat        =   "00"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D Back_Bin 
                  Height          =   420
                  Index           =   12
                  Left            =   8550
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   8
                  Begin VB.CheckBox ChkB1_ZHKbn 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Check1"
                     Height          =   345
                     Index           =   12
                     Left            =   390
                     TabIndex        =   139
                     Top             =   30
                     Visible         =   0   'False
                     Width           =   180
                  End
                  Begin CisText_V60.CisText iB1_Bin 
                     Height          =   360
                     Index           =   12
                     Left            =   30
                     TabIndex        =   138
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   635
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
                     cFdAutoFormat   =   1
                     cGFormat        =   "##"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D Back_Bin 
                  Height          =   420
                  Index           =   13
                  Left            =   9165
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   8
                  Begin VB.CheckBox ChkB1_ZHKbn 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Check1"
                     Height          =   345
                     Index           =   13
                     Left            =   390
                     TabIndex        =   141
                     Top             =   30
                     Visible         =   0   'False
                     Width           =   210
                  End
                  Begin CisText_V60.CisText iB1_Bin 
                     Height          =   360
                     Index           =   13
                     Left            =   30
                     TabIndex        =   140
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   635
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
                     cGFormat        =   "00"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D Back_Bin 
                  Height          =   420
                  Index           =   14
                  Left            =   9780
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   8
                  Begin VB.CheckBox ChkB1_ZHKbn 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Check1"
                     Height          =   345
                     Index           =   14
                     Left            =   390
                     TabIndex        =   143
                     Top             =   30
                     Visible         =   0   'False
                     Width           =   210
                  End
                  Begin CisText_V60.CisText iB1_Bin 
                     Height          =   360
                     Index           =   14
                     Left            =   30
                     TabIndex        =   142
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   635
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
                     cGFormat        =   "00"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D Back_Bin 
                  Height          =   420
                  Index           =   15
                  Left            =   10395
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   8
                  Begin VB.CheckBox ChkB1_ZHKbn 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Check1"
                     Height          =   345
                     Index           =   15
                     Left            =   390
                     TabIndex        =   145
                     Top             =   30
                     Visible         =   0   'False
                     Width           =   180
                  End
                  Begin CisText_V60.CisText iB1_Bin 
                     Height          =   360
                     Index           =   15
                     Left            =   30
                     TabIndex        =   144
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   635
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
                     cFdAutoFormat   =   1
                     cGFormat        =   "##"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D Back_Bin 
                  Height          =   420
                  Index           =   16
                  Left            =   11010
                  Top             =   0
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   8
                  Begin VB.CheckBox ChkB1_ZHKbn 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "Check1"
                     Height          =   345
                     Index           =   16
                     Left            =   390
                     TabIndex        =   147
                     Top             =   30
                     Visible         =   0   'False
                     Width           =   210
                  End
                  Begin CisText_V60.CisText iB1_Bin 
                     Height          =   360
                     Index           =   16
                     Left            =   30
                     TabIndex        =   146
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   635
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
                     cGFormat        =   "00"
                     cILength        =   2
                  End
               End
            End
            Begin Cis3D_v60.CIS3D DispTorcd 
               Height          =   435
               Left            =   870
               Top             =   840
               Visible         =   0   'False
               Width           =   3900
               _ExtentX        =   6879
               _ExtentY        =   767
               BackColor       =   16761087
               Caption         =   "XXXXXXX"
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
               cAlingnment     =   1
               cPositionX      =   60
            End
         End
         Begin CisText_V60.CisText B_TabDummy 
            Height          =   435
            Left            =   180
            TabIndex        =   148
            Top             =   1215
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   767
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin CisText_V60.CisText B_TabDummy2 
            Height          =   420
            Left            =   165
            TabIndex        =   149
            Top             =   720
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   741
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "Åy î≠íçÉeÅ[ÉuÉãÅ@ÉÅÉìÉe  Åz"
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
         cBoderColor2    =   12632256
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14340
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
         Left            =   11790
         Top             =   30
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
End
Attribute VB_Name = "CKN0120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :   î≠íçÉeÅ[ÉuÉãÉÅÉìÉeÉiÉìÉX
'**       ÉtÉHÅ[ÉÄID    :   CKN0120
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :   2005/02/28  By CIS SIMPR-A
'**       ïœ  çX  ì˙    :   2005/04/19  ëOì˙î≠çsãÊï™ÇÃí«â¡
'**       ïœ  çX  ì˙    :   2005/05/09  î≠íçéûä‘ÇïœçXÇ≥ÇÍÇΩéûëSçsÇçXêVÇ∑ÇÈ
'**       ïœ  çX  ì˙    :   2006/03/29  ∏ﬁÿØƒﬁèÓïÒÇ…édì¸êÊ∫∞ƒﬁ+ñºèÃÇï\é¶Ç∑ÇÈ(âE∏ÿØ∏)
'**                                     ∏ﬁÿØƒﬁå©èoÇµÇ…ÇƒâE∏ÿØ∏ÇçsÇ§Ç∆¥◊∞Ç…Ç»ÇÈÇÃÇèCê≥
'**       ïœ  çX  ì˙    :   2007/06/04  àÍóóï\ÇÃç≈èIçsãyÇ—ÅAéüï≈êÊì™çsÇ™Ç®Ç©ÇµÇ¢
'**       ïœ  çX  ì˙    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    
    Dim FormAct         As Boolean      ' Ã´∞—±∏√®Ãﬁ Ã◊∏ﬁ
    Dim ProcHB          As String * 2   ' ì¸óÕàÊêÿë÷
    Dim SyoriKB         As String       ' èàóùãÊï™(êVãK•èCê≥•çÌèú ìô)
    Dim SyoriNM         As String       ' èàóùñºèÃ(êVãK•èCê≥•çÌèú ìô)

    Dim ActObj          As Object
    Dim ErrObj          As Object
    Dim ActRow          As Integer
    Dim cMouseRow       As Integer
    Dim cMouseCol       As Integer
    Dim MaxKaisu        As Byte         'ÉVÉXÉeÉÄç≈ëÂâÒêî
    Dim mMaxKaisu       As Byte         'ê›íËç≈ëÂâÒêî
    Dim mHCTime()       As String       'î≠íçéûä‘ëﬁî
    Dim mRETime()       As String       'âÒé˚éûä‘ëﬁî
    Dim mHKTime()       As String       'à¯éÊéûä‘ëﬁî
    Dim mHYKbn()        As Byte         'íãñÈãÊï™
    Dim HTBinMax        As Byte         'ïiî‘éÊà¯êÊÉ}ÉXÉ^Å@ÉTÉCÉNÉã(âÒ)ç≈ëÂ
    Dim mBinTime()      As String       'ï÷ê›íËéûä‘
    Dim mBinKai()       As Byte         'ï÷ÇÃâÒêîëﬁî
    
    
    Dim BinMax          As Byte         'ÉTÉCÉNÉãÅiâÒÅj
    Dim BinDay          As Byte         'ÉTÉCÉNÉãÅiì˙Åj
    Dim mTorcd          As String       'édì¸êÊëﬁî
    Dim mUkeir          As String       'éÛì¸ëﬁî
    Dim mCycle          As String       'ÉTÉCÉNÉãëﬁî
    Dim mEndHYmd        As String       'ç≈èIî≠íçì˙ëﬁî
    Dim mEndHBin        As String       'ç≈èIî≠íçï÷ëﬁî
    Dim mEndHKai        As Byte         'ç≈èIî≠íçèàóùâÒêîëﬁî
    
    Dim mFTime          As String       'êUèoéûä‘ëﬁî
    Dim mHTime          As String       'î≠çséûä‘ëﬁî
    Dim mKTime          As String       'âÒé˚éûä‘ëﬁî
    Dim mNTime          As String       'ì¸â◊éûä‘ëﬁî
    Dim mSTTime()       As String       'äJénéûä‘
    Dim mEDTime()       As String       'èIóπéûä‘
    Dim mFurBin()       As Byte         'êUèoï÷
    Dim mEdaNo()        As Byte         'çséØï 
    Dim mHCDay()        As Byte         'î≠íçâ¡éZì˙
    Dim mHCBin()        As Byte         'î≠íçï÷
    Dim mFurDay()       As Integer      'êUèoâ¡éZì˙
    Dim mNYDay()        As Byte         'î[ì¸â¡éZì˙
    
    Dim mPrtDay()       As Byte         'î≠çsâ¡éZì˙ÅiëOì˙î≠çsÅj
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
    
    '+-----------------------+
    '+  ÉOÉäÉbÉhèâä˙ê›íË
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = vsGrid
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 430
        .PatanMax = 1
        .SelectCol = 59
        .InitGet = False
        .Init
    End With
    Call GridInit
'   #------------------#
'   #    èâä˙ì¸óÕê›íË   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
    SyoriKB = "KAN"
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
'
    MaxKaisu = 16: mMaxKaisu = 16

' î[ì¸ì˙ÅEï÷éZèoï˚ñ@(ÉVÉXÉeÉÄê›íËíl)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 14
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 14
        IRN.oyian = "0"
    End If
    If CisFun.Val2(Left(IRN.oyian, 1)) = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    î[ì¸ì˙•ï÷ÇÃéZèoï˚ñ@ÇÕ [éûä‘] ÇÃê›íËÇ…Ç»Ç¡ÇƒÇ¢Ç‹Ç∑ÅI            "
            .MB_MSG(4) = "    î≠íçÉTÉCÉNÉãÉ}ÉXÉ^ÉÅÉìÉeÇ≈éûä‘ÇÃê›íËÇçsÇ¡Çƒâ∫Ç≥Ç¢ÅB            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
    
    End If
    
    Call HeadSet
    Call BodySet
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
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    'Åyí èÌèCê≥Åz
           Case vbKeyF5:       If PB_KAN.Visible Then Call PB_KAN_Click    'Åyä»à’èCê≥Åz
           Case vbKeyF8:       If PB_PRT.Visible Then Call PB_PRT_Click    'ÅyàÛç¸Åz
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    'ÅyéÊè¡Åz
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    'ÅyèIóπÅz
           Case vbKeyReturn
                               If PB_Ent.Visible Then                      'Åyì¸óÕÅz
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
    '
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
                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
   Call ReturnPress
                                  Dummy.Enabled = False
End Sub
'+----------------------------+
'+    éÊ  è¡  ÉL Å[(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
   Call CisFrm.MousePT(11)
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
   
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call SyoriName(SyoriKB)
      Call HeadBodyClear("B")
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    àÛÅ@ç¸  ÉL Å[(F8)       +
'+----------------------------+
Private Sub PB_PRT_Click()
    CKN0125_MaxKaisu = CisFun.Val2(H1lb_MaxKai)
'   ÉäÉXÉgèoóÕêÊämîF
    CKN0125.Show vbModal
    Unload CKN0125
    Set CKN0125 = Nothing
''   àÛç¸èàóù
    If CKN0125_PrintOut <> 9 Then
        Call CKN0125LPrint(CKN0125_PrintOut)
    End If
End Sub
'+----------------------------+
'+    í èÌèCê≥  ÉL Å[(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If SyoriKB <> "MNT" Then
        SyoriKB = "MNT"
    End If
    Call SyoriName(SyoriKB)
End Sub
'+----------------------------+
'+    ä»à’èCê≥  ÉL Å[(F5)       +
'+----------------------------+
Private Sub PB_KAN_Click()
    If SyoriKB <> "KAN" Then
        SyoriKB = "KAN"
    End If
    Call SyoriName(SyoriKB)
End Sub
'+-----------------------------+
'+  ChkH1_HYKb
'+-----------------------------+
Private Sub ChkH1_HYKb_Click(Index As Integer)
    If ChkH1_HYKb(Index).Value Then
        ChkH1_HYKb(Index).Caption = "ñÈ"
    Else
        ChkH1_HYKb(Index).Caption = "íã"
    End If
End Sub
'+-----------------------------+
'+  iB1_EHYmd
'+-----------------------------+
Private Sub iB1_EHYmd_GotFocus()
    Set ActObj = iB1_EHYmd
    iB1_Bin(8).Tag = ""
End Sub
Private Sub iB1_EHYmd_KeyDown(KeyCode As Integer, Shift As Integer)
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 26) = "M"
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_EHYmd_LostFocus()
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 21) = iB1_EHYmd
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 26) = "M"
    If RTrim(iB1_EHYmd) <> "" Then
        If iB1_EHYmd.cChkResult Then
            With CisFun
                .DateE_BefVal = RTrim(iB1_EHYmd)
                .DateEdit Slash
                vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 4) = .DateE_AfVal
            End With
        Else
            vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 4) = RTrim(iB1_EHYmd)
        End If
    Else
        vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 4) = ""
    End If
    If CisFun.Val2(iB1_EHBin) <> 0 Then
        vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 4) = vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 4) & "-" & Format(CisFun.Val2(iB1_EHBin), "00")
    Else
        vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 4) = ""
    End If
    
End Sub
'+-----------------------------+
'+  iB1_EHBin
'+-----------------------------+
Private Sub iB1_EHBin_GotFocus()
    Set ActObj = iB1_EHBin
End Sub
Private Sub iB1_EHBin_KeyDown(KeyCode As Integer, Shift As Integer)
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 26) = "M"
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_EHBin_LostFocus()
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 22) = iB1_EHBin
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 26) = "M"
    If RTrim(iB1_EHYmd) <> "" Then
        With CisFun
            .DateE_BefVal = RTrim(iB1_EHYmd)
            .DateEdit Slash
            vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 4) = .DateE_AfVal
        End With
    Else
        vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 4) = ""
    End If
    If CisFun.Val2(iB1_EHBin) <> 0 Then
        vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 4) = vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 4) & "-" & Format(CisFun.Val2(iB1_EHBin), "00")
    Else
        vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 4) = ""
    End If
End Sub
'+-----------------------------+
'+  iB1_Bin
'+-----------------------------+
Private Sub iB1_Bin_GotFocus(Index As Integer)
    Set ActObj = iB1_Bin
    iB1_Bin(1).Tag = Format(Index, "0")
'''    If Index = 1 Then
'''        iB1_Bin(8).Tag = ""
'''    End If
'''    If Index = mMaxKaisu Then
'''        iB1_Bin(8).Tag = "E"
'''    End If
End Sub
Private Sub iB1_Bin_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 26) = "M"
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Bin_LostFocus(Index As Integer)
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), Index + 26) = iB1_Bin(Index)
    If CisFun.Val2(iB1_Bin(Index)) <> 0 And CisFun.Val2(vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), Index + 42)) <> 0 Then
        vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), Index + 4) = iB1_Bin(Index) & "*"
    Else
        vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), Index + 4) = iB1_Bin(Index)
    End If
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 26) = "M"
End Sub
'+-----------------------------+
'+  ChkB1_ZHKbn
'+-----------------------------+
Private Sub ChkB1_ZHKbn_GotFocus(Index As Integer)
    Set ActObj = ChkB1_ZHKbn
    ChkB1_ZHKbn(1).Tag = Format(Index, "0")
    If Index = mMaxKaisu Then
        iB1_Bin(8).Tag = "E"
    End If
End Sub
Private Sub ChkB1_ZHKbn_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub ChkB1_ZHKbn_LostFocus(Index As Integer)
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), Index + 42) = ChkB1_ZHKbn(Index).Value
    If CisFun.Val2(vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), Index + 26)) <> 0 And CisFun.Val2(vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), Index + 42)) <> 0 Then
        vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), Index + 4) = vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), Index + 26) & "*"
    Else
        vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), Index + 4) = vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), Index + 26)
    End If
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 26) = "M"
End Sub
'+-----------------------------+
'+  B_TabDummy2
'+-----------------------------+
Private Sub B_TabDummy2_GotFocus()
    If iB1_Bin(8).Tag = "" Then
        Call TextToVsGrid(Val(Input_Row.Tag))
        Call Grid_RowChange(vbKeyUp)
        iB1_Bin(mMaxKaisu).SetFocus
        Exit Sub
    Else
        If Val(Input_Row.Tag) = 1 Then
           iB1_EHYmd.SetFocus
        Else
           Call TextToVsGrid(Val(Input_Row.Tag))
           Call Grid_RowChange(vbKeyUp)
           iB1_EHYmd.SetFocus
        End If
    End If
End Sub
'+-----------------------------+
'+  B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
    If iB1_Bin(8).Tag = "S" Then
        Call TextToVsGrid(Val(Input_Row.Tag))
        Call Grid_RowChange(vbKeyUp)
    End If
    If iB1_Bin(8).Tag = "E" Then
        Call TextToVsGrid(Val(Input_Row.Tag))
        Call Grid_RowChange(vbKeyDown)
    End If
    'SendKeys "{Tab}"
    iB1_EHYmd.SetFocus
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
    B1_Area1.Enabled = B1Mode Or H1Mode
    B1_Area1.BackColor = B1Color
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    PB_PRT.Visible = H1Mode
    PB_MNT.Visible = H1Mode
    PB_KAN.Visible = H1Mode

    Call GridTextEnabled(B1Mode)
    
' ¿ﬁ–∞çÄñ⁄(ì¸óÕã÷é~)
    Dummy.Enabled = False
'
End Function
'**************************************
'*                                    *
'*      èàóùãÊï™Ç…ÇÊÇËãÊï™ñºï\é¶        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   Dim w_Color     As Long
   w_Color = &H80000008 'èâä˙êFæØƒ(çï)
   Select Case SyoriKB
     Case "MNT"
          SyoriNM = "í èÌ"
          K_Sykbnm.ForeColor = &HFF0000  '(ê¬êF)
     Case "KAN"
          SyoriNM = "ä»à’"
          K_Sykbnm.ForeColor = &HFF00FF   '(ÉsÉìÉN)
          PB_MNT.Visible = False
          PB_KAN.Visible = False
     Case Else
          SyoriNM = ""
   End Select
   For Each gObj In Me.Controls
       If gObj.Name Like "i*" Then gObj.ForeColor = w_Color
       If gObj.Name Like "*Comb*" Then gObj.ForeColor = w_Color
       If gObj.Name Like "*vsGrid*" Then gObj.ForeColor = w_Color
'      * PF-Key ï\é¶ê›íË *
       If SyoriKB <> "LOC" Then
          If gObj.Name = "PB_KAN" Or gObj.Name = "PB_MNT" Then
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
                If .Name Like "iH*" Then
                    .Enabled = True
                End If
            End If
            If .Name Like "Chk" & wbc_Enm & "_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
        For gInt = 1 To MaxKaisu
            If gInt <= mMaxKaisu Then
                iB1_Bin(gInt).Visible = True
                ChkB1_ZHKbn(gInt).Visible = True
            Else
                iB1_Bin(gInt).Visible = False
                ChkB1_ZHKbn(gInt).Visible = True
            End If
        Next
        CisVsGrid.Clear
    End If
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
        vsGrid.Row = 1
        Call VsGridToText(vsGrid.Row)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        ProcHB = "H1"
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
    Head1Chk = False
'    Dim YMD     As Date
    Dim wCol    As Byte
    Dim wBIdx   As Byte
    Dim w24h    As Byte
    Dim wHYKb   As Byte
    Dim wHYCnt  As Byte
'<< ç≈èIî≠íçâÒÇãÅÇﬂÇÈ >>
    For gCnt1 = 1 To MaxKaisu
        If RTrim(iH1_HHh(gCnt1)) <> "" Or RTrim(iH1_HMm(gCnt1)) <> "" Then
            mMaxKaisu = gCnt1
        End If
    Next gCnt1
    H1lb_MaxKai = mMaxKaisu
    If mMaxKaisu < HTBinMax Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    î≠íçâÒêî < ÉTÉCÉNÉã(âÒ)            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Set ErrObj = iH1_HHh(1)
        GoTo Head1Chk_Err
    End If
'<< ÉèÅ[ÉNèâä˙ê›íË >>
    wHYKb = ChkH1_HYKb(1).Value: wHYCnt = 0: mFTime = Space(MaxKaisu * 4)
    Erase mHCTime: Erase mRETime: Erase mHKTime: Erase mHYKbn
    ReDim mHCTime(MaxKaisu): ReDim mRETime(MaxKaisu): ReDim mHKTime(MaxKaisu): ReDim mHYKbn(MaxKaisu)
    ReDim mBinTime(MaxKaisu)
    wCol = 1:  w24h = 0: wBIdx = 0
    
    For gCnt1 = 1 To mMaxKaisu
        mBinTime(gCnt1) = ""
'<< î≠íçéûä‘ >>É`ÉFÉbÉN
        If CisFun.Val2(iH1_HHh(gCnt1)) <> 0 Then
            If CisFun.Val2(iH1_HHh(gCnt1)) >= 24 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    î≠íçéûä‘•éû çƒì¸óÕ            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iH1_HHh(gCnt1)
                GoTo Head1Chk_Err
            End If
        End If
        If CisFun.Val2(iH1_HMm(gCnt1)) <> 0 Then
            If CisFun.Val2(iH1_HMm(gCnt1)) >= 60 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    î≠íçéûä‘•ï™ çƒì¸óÕ            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iH1_HMm(gCnt1)
                GoTo Head1Chk_Err
            End If
        End If
        iH1_HHh(gCnt1) = Format(CisFun.Val2(iH1_HHh(gCnt1)), "00")
        iH1_HMm(gCnt1) = Format(CisFun.Val2(iH1_HMm(gCnt1)), "00")
        mHCTime(gCnt1) = iH1_HHh((gCnt1)) & iH1_HMm(gCnt1)
'<< âÒé˚éûä‘ >>É`ÉFÉbÉN
        If CisFun.Val2(iH1_KHh(gCnt1)) <> 0 Then
            If CisFun.Val2(iH1_KHh(gCnt1)) >= 24 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    âÒé˚éûä‘•éû çƒì¸óÕ            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iH1_KHh(gCnt1)
                GoTo Head1Chk_Err
            End If
        End If
        If CisFun.Val2(iH1_KMm(gCnt1)) <> 0 Then
            If CisFun.Val2(iH1_KMm(gCnt1)) >= 60 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    âÒé˚éûä‘•ï™ çƒì¸óÕ            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iH1_KMm(gCnt1)
                GoTo Head1Chk_Err
            End If
        End If
'        iH1_KHh(gCnt1) = Format(CisFun.Val2(iH1_KHh(gCnt1)), "00")
'        iH1_KMm(gCnt1) = Format(CisFun.Val2(iH1_KMm(gCnt1)), "00")
        mRETime(gCnt1) = iH1_KHh((gCnt1)) & iH1_KMm(gCnt1)
'<< à¯éÊéûä‘ >>É`ÉFÉbÉN
        If CisFun.Val2(iH1_THh(gCnt1)) <> 0 Then
            If CisFun.Val2(iH1_THh(gCnt1)) >= 24 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    à¯éÊéûä‘•éû çƒì¸óÕ            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iH1_THh(gCnt1)
                GoTo Head1Chk_Err
            End If
        End If
        If CisFun.Val2(iH1_TMm(gCnt1)) <> 0 Then
            If CisFun.Val2(iH1_TMm(gCnt1)) >= 60 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    à¯éÊéûä‘•ï™ çƒì¸óÕ            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iH1_TMm(gCnt1)
                GoTo Head1Chk_Err
            End If
        End If
'        iH1_THh(gCnt1) = Format(CisFun.Val2(iH1_THh(gCnt1)), "00")
'        iH1_TMm(gCnt1) = Format(CisFun.Val2(iH1_TMm(gCnt1)), "00")
        mHKTime(gCnt1) = iH1_THh((gCnt1)) & iH1_TMm(gCnt1)
        
'<< íãñÈãŒãÊï™ >>É`ÉFÉbÉN
        If wHYKb <> ChkH1_HYKb(gCnt1).Value Then
            wHYCnt = wHYCnt + 1
            If wHYCnt > 2 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    íãñÈãŒãÊï™Å@ê›íËàŸèÌÅI            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = ChkH1_HYKb(gCnt1)
                GoTo Head1Chk_Err
            End If
            wHYKb = ChkH1_HYKb(gCnt1).Value
        End If
        mHYKbn(gCnt1) = ChkH1_HYKb(gCnt1).Value
'
        Mid(mFTime, wCol, 4) = mHCTime(gCnt1)
        mBinTime(gCnt1) = mHCTime(gCnt1)
'êUèoéûä‘ê›íËÅ@îÕàÕèdï°É`ÉFÉbÉN
        wBIdx = wBIdx + 1
        If wBIdx <> 1 Then
            If mBinTime(gCnt1) <= mBinTime(gCnt1 - 1) Then
                If w24h >= 1 Then
                    With CisFun
                        .MB_Lines = 4
                        .MB_MSG(2) = "    Å@éûä‘ê›íËÉGÉâÅ[Åiéûä‘èdï°Åj            "
                        .MB_Title = ""
                        .MB_Button = Error
                        .MBOX
                    End With
                    Set ErrObj = iH1_HHh(gCnt1)
                    GoTo Head1Chk_Err
                End If
                w24h = w24h + 2
            End If
            For gCnt2 = 1 To wBIdx - 1
                If mBinTime(gCnt1) = mBinTime(gCnt2) Then
                    With CisFun
                        .MB_Lines = 4
                        .MB_MSG(2) = "    Å@éûä‘ê›íËÉGÉâÅ[ÅiìØàÍéûä‘óLÇËÅj            "
                        .MB_Title = ""
                        .MB_Button = Error
                        .MBOX
                    End With
                    Set ErrObj = iH1_HHh(gCnt1)
                    GoTo Head1Chk_Err
                End If
            Next
            If mBinTime(gCnt1) >= mBinTime(1) Then
                If w24h >= 2 Then
                    With CisFun
                        .MB_Lines = 4
                        .MB_MSG(2) = "      éûä‘ê›íËÉGÉâÅ[Åiéûä‘èdï°Åj            "
                        .MB_Title = ""
                        .MB_Button = Error
                        .MBOX
                    End With
                    Set ErrObj = iH1_HHh(gCnt1)
                    GoTo Head1Chk_Err
                End If
            End If
        Else
            If mBinTime(gCnt1) = "0000" Then
                w24h = w24h + 1
            End If
        End If
'
        wCol = wCol + 4
    Next gCnt1

    If mHYKbn(1) <> 0 Then
        If mBinTime(1) > mBinTime(mMaxKaisu) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "      ÇPï÷Ç™ñÈãŒÇÃéûÅA1ï÷î≠íçéûä‘Å@>Å@ç≈èIï÷î≠íçéûä‘            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            Set ErrObj = iH1_HHh(1)
            gCnt1 = 1
            GoTo Head1Chk_Err
        End If
    End If
       
    Call BodySet
    Head1Chk = True
    Exit Function
Head1Chk_Err:
    ErrObj.SetFocus
End Function
'****************************
'*      Body1¡™Ø∏èàóù
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
    Dim wSTCol  As Byte
    Dim wEDCol  As Byte
    Dim wCycleBin    As Byte
    Dim wBinSet()    As Byte
    Dim wZenPrt()    As Byte
    Dim wZenPKb    As Byte
    Dim wGyakuKb   As Byte
    Dim wBinCnt    As Byte
    
    If SyoriKB = "MNT" Then
        Call TextToVsGrid(Val(Input_Row.Tag))
        Call GridTextEnabled(False)
    End If
    
    wSTCol = 27: wEDCol = wSTCol + mMaxKaisu - 1
    For gCnt1 = 1 To CisFun.Val2(iT1_Kensu)
        wCycleBin = CisFun.Val2(Mid(vsGrid.TextMatrix(gCnt1, 24), 2, 2))
        'ç≈èIî≠íçì˙
        If RTrim(vsGrid.TextMatrix(gCnt1, 21)) <> "" Then
            iB1_EHYmd = RTrim(vsGrid.TextMatrix(gCnt1, 21))
            If Not iB1_EHYmd.cChkResult Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    ç≈èIî≠íçì˙Å@àŸèÌÅI            "
                    .MB_MSG(4) = "    Å@Å@" & gCnt1 & "Å@çsñ⁄            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iB1_EHYmd
                GoTo Body1Chk_Err
            End If
        End If
        'ç≈èIî≠íçï÷
        If CisFun.Val2(vsGrid.TextMatrix(gCnt1, 22)) <> 0 Then
            If RTrim(vsGrid.TextMatrix(gCnt1, 21)) = "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    ç≈èIî≠íçì˙Å@ïKê{ÅI            "
                    .MB_MSG(4) = "    Å@Å@" & gCnt1 & "Å@çsñ⁄            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iB1_EHYmd
                GoTo Body1Chk_Err
            End If
            If CisFun.Val2(vsGrid.TextMatrix(gCnt1, 22)) > wCycleBin Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    ç≈èIî≠íçï÷Å@ÅÑÅ@ÉTÉCÉNÉãÅiâÒÅjÅI            "
                    .MB_MSG(4) = "    Å@Å@" & gCnt1 & "Å@çsñ⁄            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iB1_EHBin
                GoTo Body1Chk_Err
            End If
        Else
            If RTrim(vsGrid.TextMatrix(gCnt1, 21)) <> "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    ç≈èIî≠íçï÷Å@ïKê{ÅI            "
                    .MB_MSG(4) = "    Å@Å@" & gCnt1 & "Å@çsñ⁄            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iB1_EHBin
                GoTo Body1Chk_Err
            End If
        End If
               
        
        Erase wBinSet: ReDim wBinSet(wCycleBin)
        For gCnt2 = 1 To wCycleBin
            For gCnt3 = wSTCol To wEDCol
                If gCnt2 = CisFun.Val2(vsGrid.TextMatrix(gCnt1, gCnt3)) Then
                    wBinSet(gCnt2) = wBinSet(gCnt2) + 1
                End If
                If CisFun.Val2(vsGrid.TextMatrix(gCnt1, gCnt3)) > wCycleBin Then
                    With CisFun
                        .MB_Lines = 4
                        .MB_MSG(2) = "    ÉTÉCÉNÉã(âÒ)ÅÉï÷Å@ÇÕïsâ¬ÅI            "
                        .MB_MSG(4) = "    Å@Å@" & gCnt1 & "Å@çsñ⁄            "
                        .MB_Title = ""
                        .MB_Button = Error
                        .MBOX
                    End With
                    Set ErrObj = iB1_Bin(1)
                    GoTo Body1Chk_Err
                End If
            Next gCnt3
        Next gCnt2
        
        wBinCnt = 0
        For gCnt2 = 1 To wCycleBin
            If wBinSet(gCnt2) > 1 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    ï÷ÇÃê›íËÇ™èdï°ÇµÇƒÇ¢Ç‹Ç∑ÅI            "
                    .MB_MSG(4) = "    Å@Å@" & gCnt1 & "Å@çsñ⁄            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iB1_Bin(1)
                GoTo Body1Chk_Err
            End If
            If wBinSet(gCnt2) > 0 Then
                wBinCnt = wBinCnt + 1
            End If
        Next gCnt2
        If wBinCnt <> 0 And wCycleBin <> wBinCnt Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ï÷ÇÃê›íËÇ™Ç†ÇËÇ‹ÇπÇÒÅI            "
                .MB_MSG(4) = "    Å@Å@" & gCnt1 & "Å@çsñ⁄            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            Set ErrObj = iB1_Bin(1)
            GoTo Body1Chk_Err
        End If
        If wBinCnt = 0 And RTrim(vsGrid.TextMatrix(gCnt1, 21)) <> "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ï÷ÇÃê›íËÇ™Ç†ÇËÇ‹ÇπÇÒÅI            "
                .MB_MSG(4) = "    Å@Å@" & gCnt1 & "Å@çsñ⁄            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            Set ErrObj = iB1_Bin(1)
            GoTo Body1Chk_Err
        End If

        gIndex = 0: gCnt2 = 0: wBinSet(0) = 0: gLong = 0: wZenPKb = 0: wGyakuKb = 0
        Erase wZenPrt: ReDim wZenPrt(wCycleBin)
        For gCnt3 = wSTCol To wEDCol
            If CisFun.Val2(vsGrid.TextMatrix(gCnt1, gCnt3)) <> 0 Then
                gLong = gLong + 1
                wZenPrt(gLong) = CisFun.Val2(vsGrid.TextMatrix(gCnt1, gCnt3 + 16))
                If CisFun.Val2(vsGrid.TextMatrix(gCnt1, gCnt3 + 16)) <> 0 Then
                    wZenPKb = 1
                End If
                If wBinSet(0) <> 0 And CisFun.Val2(vsGrid.TextMatrix(gCnt1, gCnt3)) < wBinSet(0) Then
                    gCnt2 = gCnt2 + 1
                    wGyakuKb = gLong
                End If
                If wBinSet(0) <> 0 And wBinSet(0) + 1 <> CisFun.Val2(vsGrid.TextMatrix(gCnt1, gCnt3)) And _
                    CisFun.Val2(vsGrid.TextMatrix(gCnt1, gCnt3)) <> 1 Then
                    gIndex = gIndex + 1
                End If
                vsGrid.TextMatrix(gCnt1, 25) = CisFun.Val2(vsGrid.TextMatrix(gCnt1, gCnt3))
                wBinSet(0) = CisFun.Val2(vsGrid.TextMatrix(gCnt1, gCnt3))
            End If
        Next gCnt3
        If gIndex <> 0 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ï÷ê›íË(ï÷èáèò)Å@àŸèÌÅI            "
                .MB_MSG(4) = "    Å@Å@" & gCnt1 & "Å@çsñ⁄            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            Set ErrObj = iB1_Bin(1)
            GoTo Body1Chk_Err
        End If
        If gCnt2 > 1 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ï÷ê›íË(ï÷èáèò)Å@àŸèÌÅI            "
                .MB_MSG(4) = "    Å@Å@" & gCnt1 & "Å@çsñ⁄            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            Set ErrObj = iB1_Bin(1)
            GoTo Body1Chk_Err
        End If
        If wZenPKb <> 0 Then
            If wGyakuKb = 0 Then
                'ê≥èáÇÕëSÇƒÉ`ÉFÉbÉN
                gLong = 0
                For gCnt3 = 1 To wCycleBin
                    If wZenPrt(gCnt3) <> 0 Then
                        gLong = gLong + 1
                    End If
                Next gCnt3
                If gLong <> wCycleBin Then
                    With CisFun
                        .MB_Lines = 4
                        .MB_MSG(2) = "    Å@Å@ëOì˙î≠çsãÊï™Å@ê›íËàŸèÌÅI            "
                        .MB_MSG(4) = "    Å@Å@" & gCnt1 & "Å@çsñ⁄            "
                        .MB_Title = ""
                        .MB_Button = Error
                        .MBOX
                    End With
                    Set ErrObj = ChkB1_ZHKbn(1)
                    GoTo Body1Chk_Err
                End If
            Else
                'ãtèáÇÕãtèáà»ç~ëSÇƒÉ`ÉFÉbÉN
                For gCnt3 = 1 To wCycleBin
                    If wGyakuKb > gCnt3 And wZenPrt(gCnt3) <> 0 Then
                        With CisFun
                            .MB_Lines = 4
                            .MB_MSG(2) = "    Å@Å@ëOì˙î≠çsãÊï™Å@ê›íËàŸèÌÅI            "
                            .MB_MSG(4) = "    Å@Å@" & gCnt1 & "Å@çsñ⁄            "
                            .MB_Title = ""
                            .MB_Button = Error
                            .MBOX
                        End With
                        Set ErrObj = ChkB1_ZHKbn(1)
                        GoTo Body1Chk_Err
                    End If
                    If wGyakuKb <= gCnt3 And wZenPrt(gCnt3) = 0 Then
                        With CisFun
                            .MB_Lines = 4
                            .MB_MSG(2) = "    Å@Å@ëOì˙î≠çsãÊï™Å@ê›íËàŸèÌÅI            "
                            .MB_MSG(4) = "    Å@Å@" & gCnt1 & "Å@çsñ⁄            "
                            .MB_Title = ""
                            .MB_Button = Error
                            .MBOX
                        End With
                        Set ErrObj = ChkB1_ZHKbn(1)
                        GoTo Body1Chk_Err
                    End If
                Next gCnt3
            End If
        End If
    Next gCnt1
            
        
    
    Body1Chk = True
    Exit Function
Body1Chk_Err:
    vsGrid.Row = gCnt1
    If SyoriKB = "MNT" Then
        Call VsGridToText(vsGrid.Row)
        Call GridTextEnabled(True)
        ErrObj.SetFocus
    Else
        vsGrid.SetFocus
    End If
End Function
'****************************
'*      êUèoï÷ê›íË
'****************************
Private Function FuriBin_Set() As Boolean
    FuriBin_Set = False
    
    Dim wBIdx As Byte
    Dim wBinIdx1 As Byte
    Dim wBinIdx2 As Byte
    Dim wGIdx As Byte
    Dim wBin  As Byte
    Dim w24h  As Byte

    Erase mSTTime: Erase mEDTime: Erase mFurBin
    Erase mEdaNo: Erase mHCDay: Erase mHCBin
    Erase mFurDay: Erase mNYDay
    ReDim mSTTime(BinMax + 2): ReDim mEDTime(BinMax + 2): ReDim mFurBin(BinMax + 2)
    ReDim mEdaNo(BinMax + 2): ReDim mHCDay(BinMax + 2): ReDim mHCBin(BinMax + 2)
    ReDim mFurDay(BinMax + 2): ReDim mNYDay(BinMax + 2)
'ëﬁî¥ÿ±∏ÿ±
    For gIndex = 1 To BinMax + 2
        mSTTime(gIndex) = ""
        mEDTime(gIndex) = ""
        mFurBin(gIndex) = 0
        mEdaNo(gIndex) = 0
        mHCDay(gIndex) = 0
        mHCBin(gIndex) = 0
        mFurDay(gIndex) = 0
        mNYDay(gIndex) = 0
    Next gIndex
    wBIdx = 0: w24h = 0
    
    For gIndex = 1 To BinMax
        If Trim(mBinTime(gIndex)) = "" Then Exit For
'äJén•èIóπ•êUèoï÷æØƒ
        wBIdx = wBIdx + 1
        If wBIdx = 1 Then
            '1ï÷
            mSTTime(wBIdx) = "0000"
            mEDTime(wBIdx) = mBinTime(gIndex)
            mSTTime(wBIdx + 1) = mBinTime(gIndex)
            mEDTime(wBIdx + 1) = "2400"
            mFurBin(wBIdx) = gIndex
            mFurBin(wBIdx + 1) = gIndex
            mEdaNo(wBIdx + 1) = 1
        Else
            If mBinTime(gIndex) >= mEDTime(wBIdx - 1) Then
                '2ï÷à»ç~ëOÇÃéûä‘ÇÊÇËëÂÇ´Ç¢éû
                mEDTime(wBIdx) = mBinTime(gIndex)
                mFurBin(wBIdx) = gIndex
                mEdaNo(wBIdx) = 0
                If w24h <> 0 Then
                    mFurDay(wBIdx) = -1
                End If
                mSTTime(wBIdx + 1) = mBinTime(gIndex)
                mEDTime(wBIdx + 1) = "2400"
                mFurBin(wBIdx + 1) = 1
                mEdaNo(wBIdx + 1) = 1
            Else
                '2ï÷à»ç~ëOÇÃéûä‘ÇÊÇËè¨Ç≥Ç¢éû
                If w24h > 0 Then Exit Function
                mFurBin(wBIdx) = gIndex
                wBIdx = wBIdx + 1
                mSTTime(wBIdx) = "0000"
                mEDTime(wBIdx) = mBinTime(gIndex)
                mSTTime(wBIdx + 1) = mBinTime(gIndex)
                mFurBin(wBIdx) = gIndex
                mEdaNo(wBIdx) = 0
                mFurDay(wBIdx) = -1
                w24h = w24h + 1
            End If
        End If
        'ç≈èIï÷ÇÃéû
        If gIndex = BinMax Then
            If mEDTime(wBIdx) < mEDTime(1) Then
                mEDTime(wBIdx + 1) = ""
                mSTTime(wBIdx + 1) = ""
                mFurBin(wBIdx + 1) = 0
                mSTTime(1) = mBinTime(gIndex)
            Else
                If w24h > 0 Then Exit Function
            End If
            If mEDTime(wBIdx + 1) = "2400" Then
                mFurDay(wBIdx + 1) = mFurDay(wBIdx + 1) + 1
            End If
        End If
    Next gIndex
    
    FuriBin_Set = True
End Function
'**************************************
'*          Headïî ì‡óeæØƒèàóù           *
'**************************************
Private Sub HeadSet()
'
' É}ÉXÉ^ì¸óÕ SELECTï∂ æØƒ
    gSL_Select = " SELECT * FROM î≠íçÉTÉCÉNÉãê›íËÉ}ÉXÉ^"
    gSL_Select = gSL_Select & " WHERE âÒêî = 0"
    If HCSRead(gSL_Select, 1) Then
        iH1_Bikou = RTrim(HCS.îıçl)
        If RTrim(HCS.ç≈èIî≠íçì˙) <> "" Then
            H1lb_EHYmd = Left(HCS.ç≈èIî≠íçì˙, 4) & "/" & Mid(HCS.ç≈èIî≠íçì˙, 5, 2) & "/" & Mid(HCS.ç≈èIî≠íçì˙, 7, 2)
            H1lb_EHKai = HCS.ç≈èIâÒ
        End If
        H1lb_MaxKai = HCS.ç≈ëÂâÒêî
    End If
        
    gSL_Select = " SELECT * FROM î≠íçÉTÉCÉNÉãê›íËÉ}ÉXÉ^"
    gSL_Select = gSL_Select & " WHERE âÒêî <> 0"
    gSL_Select = gSL_Select & " ORDER BY âÒêî"
    If HCSRead(gSL_Select, 0) Then
        Do Until Not HCS_RDSTS
            With HCS
                If RTrim(.âÒé˚éûä‘) <> "" And RTrim(.âÒé˚éûä‘) <> "0000" Then
                    iH1_KHh(.âÒêî) = Left(.âÒé˚éûä‘, 2)
                    iH1_KMm(.âÒêî) = Mid(.âÒé˚éûä‘, 3, 2)
                End If
                If RTrim(.î≠íçéûä‘) <> "" And RTrim(.î≠íçéûä‘) <> "0000" Then
                    iH1_HHh(.âÒêî) = Left(.î≠íçéûä‘, 2)
                    iH1_HMm(.âÒêî) = Mid(.î≠íçéûä‘, 3, 2)
                End If
                If RTrim(.à¯éÊéûä‘) <> "" And RTrim(.à¯éÊéûä‘) <> "0000" Then
                    iH1_THh(.âÒêî) = Left(.à¯éÊéûä‘, 2)
                    iH1_TMm(.âÒêî) = Mid(.à¯éÊéûä‘, 3, 2)
                End If
                ChkH1_HYKb(.âÒêî).Value = .íãñÈãÊï™
            End With
            Call HCSReadNext
        Loop
    End If
    Call HCSClose
End Sub
'**************************************
'*          BODY ì‡óeæØƒèàóù           *
'**************************************
Private Sub BodySet()
'
    Call HeadBodyClear("B")
' É}ÉXÉ^ì¸óÕ SELECTï∂ æØƒ
    gSL_Select = " SELECT HT.éÊà¯êÊ édì¸êÊ,HT.éÛì¸, "
    gSL_Select = gSL_Select & " SUBSTRING(HT.ÉTÉCÉNÉã,1,1) + '-' + SUBSTRING(HT.ÉTÉCÉNÉã,2,2) + '-' + SUBSTRING(HT.ÉTÉCÉNÉã,4,2) ï“èWÉTÉCÉNÉã,"
    gSL_Select = gSL_Select & " HT.ÉTÉCÉNÉã,ISNULL(TR.ó™èÃ,'')édì¸êÊñº,"
    gSL_Select = gSL_Select & " ISNULL(HC.ç≈èIî≠íçì˙,'')ç≈èIî≠íçì˙,ISNULL(HC.ç≈èIî≠íçï÷,'')ç≈èIî≠íçï÷,ISNULL(HC.êUèoï÷,'')êUèoï÷,ISNULL(HC.î≠íçèàóùâÒêî,0)î≠íçèàóùâÒêî,"
    gSL_Select = gSL_Select & " ISNULL(HC.î≠çsâ¡éZì˙,'')î≠çsâ¡éZì˙,ISNULL(TU.ójì˙î≠íç,0) ójì˙î≠íç"
    gSL_Select = gSL_Select & " FROM ïiî‘éÊà¯êÊÉ}ÉXÉ^ HT"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN î≠íçÉTÉCÉNÉãÉ}ÉXÉ^ HC"
    gSL_Select = gSL_Select & "   ON  HC.édì¸êÊ = HT.éÊà¯êÊ"
    gSL_Select = gSL_Select & "   AND HC.ÉTÉCÉNÉã = HT.ÉTÉCÉNÉã"
    gSL_Select = gSL_Select & "   AND HC.éÛì¸ = HT.éÛì¸"
    gSL_Select = gSL_Select & "   AND HC.çséØï  = 0"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN éÊà¯êÊÉ}ÉXÉ^ TR"
    gSL_Select = gSL_Select & "   ON  TR.éÊà¯êÊCD = HT.éÊà¯êÊ"
    gSL_Select = gSL_Select & "   AND TR.éÊà¯êÊãÊï™ = 1"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN éÊà¯êÊéÛì¸É}ÉXÉ^ TU"
    gSL_Select = gSL_Select & "   ON  TU.éÊà¯êÊCD = HT.éÊà¯êÊ"
    gSL_Select = gSL_Select & "   AND TU.éÊà¯êÊãÊï™ = 1"
    gSL_Select = gSL_Select & "   AND TU.éÛì¸ = HT.éÛì¸"
    gSL_Select = gSL_Select & " WHERE HT.ïiñ⁄ = 1"
    gSL_Select = gSL_Select & "   AND TR.é–ì‡ãÊï™ = 0"
    gSL_Select = gSL_Select & "   AND HT.ÉTÉCÉNÉã <> '' AND HT.ÉTÉCÉNÉã <> '00000'"
    gSL_Select = gSL_Select & "   AND SUBSTRING(HT.ÉTÉCÉNÉã,2,2) <= '16'"
'    gSL_Select = gSL_Select & "   AND SUBSTRING(HT.ÉTÉCÉNÉã,2,2) <= '" & Format(MaxKaisu, "00") & "'"
    gSL_Select = gSL_Select & " GROUP BY HT.éÊà¯êÊ,HT.éÛì¸,HT.ÉTÉCÉNÉã,TR.ó™èÃ,HC.êUèoï÷,HC.î≠íçèàóùâÒêî,HC.ç≈èIî≠íçì˙,HC.ç≈èIî≠íçï÷,HC.î≠çsâ¡éZì˙,TU.ójì˙î≠íç"
    gSL_Select = gSL_Select & " ORDER BY HT.éÊà¯êÊ,HT.éÛì¸,HT.ÉTÉCÉNÉã,TR.ó™èÃ,HC.êUèoï÷,HC.î≠íçèàóùâÒêî,HC.ç≈èIî≠íçì˙,HC.ç≈èIî≠íçï÷,HC.î≠çsâ¡éZì˙"
    If HCMRead(gSL_Select, 0) Then
        Call GridSet
        Call HCMClose
    End If
    Call HCMClose
    
    gSL_Select = " SELECT MAX(SUBSTRING(ÉTÉCÉNÉã,2,2)) ÉTÉCÉNÉã FROM ïiî‘éÊà¯êÊÉ}ÉXÉ^"
    gSL_Select = gSL_Select & " WHERE ïiñ⁄ = 1"
    gSL_Select = gSL_Select & "   AND ÉTÉCÉNÉã <> ''"
    If HCMRead(gSL_Select, 1) Then
        If MaxKaisu >= CisFun.Val2(HCM.ÉTÉCÉNÉã) Then
            HTBinMax = CisFun.Val2(HCM.ÉTÉCÉNÉã)
        Else
            HTBinMax = MaxKaisu
        End If
    Else
        HTBinMax = 0
    End If
    If CisFun.Val2(H1lb_MaxKai) = 0 Then
        H1lb_MaxKai = HTBinMax
    End If
    
End Sub
'+----------------------+
'+    í«â¡•çXêV èàóù     +
'+----------------------+
Private Function DBPut()
    DBPut = False
    Dim AddFlg  As Boolean
    Dim HCSCHGFlg  As Boolean
    Dim wYDJun As Byte

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     çXêVèàóùÇçsÇ¢Ç‹Ç∑      "
        .MB_Title = "çXêVèàóù"
        .MB_MSG(3) = "     ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH      "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            If SyoriKB = "MNT" Then
                Call VsGridToText(Val(Input_Row.Tag))
                Call GridTextEnabled(True)
            Else
                vsGrid.SetFocus
            End If
            Exit Function
        End If
    End With
    On Error GoTo DBPut_Err
    PNL_Wait.Visible = True
    DoEvents
    
    Call CisDB.DBTran(TransBegin)  ' ƒ◊›ªﬁ∏ºÆ›äJén
'-----------------------------------'
'<< î≠íçÉTÉCÉNÉãê›íËÉ}ÉXÉ^ çXêVÅ@>>
'-----------------------------------'
    HCSCHGFlg = False
    With CisDB
        On Error GoTo DBPut_Err
' ( çÌèúèàóù ) î≠íçÉTÉCÉNÉãê›íËÉ}ÉXÉ^
        .SQL = "DELETE FROM î≠íçÉTÉCÉNÉãê›íËÉ}ÉXÉ^ "
        .SQL = .SQL & " WHERE âÒêî > " & mMaxKaisu
        .DBExec
    End With
    For gInt = 0 To mMaxKaisu
        gSL_Select = "SELECT * FROM î≠íçÉTÉCÉNÉãê›íËÉ}ÉXÉ^"
        gSL_Select = gSL_Select & " WHERE âÒêî = " & gInt
        If HCSRead(gSL_Select, 1) Then
            AddFlg = False
            If gInt <> 0 Then
                If HCS.î≠íçéûä‘ <> iH1_HHh(gInt) & iH1_HMm(gInt) Then
                    HCSCHGFlg = True
                End If
            End If
        Else
            AddFlg = True
        End If
        Call ItemsClearHCS
        With HCS
            .âÒêî = gInt
            If gInt <> 0 Then
                .íãñÈãÊï™ = ChkH1_HYKb(gInt).Value
                .âÒé˚éûä‘ = Format(CisFun.Val2(iH1_KHh(gInt)), "00") & Format(CisFun.Val2(iH1_KMm(gInt)), "00")
                .î≠íçéûä‘ = Format(CisFun.Val2(iH1_HHh(gInt)), "00") & Format(CisFun.Val2(iH1_HMm(gInt)), "00")
                .à¯éÊéûä‘ = Format(CisFun.Val2(iH1_THh(gInt)), "00") & Format(CisFun.Val2(iH1_TMm(gInt)), "00")
            Else
                .îıçl = RTrim(iH1_Bikou)
            End If
            .ç≈ëÂâÒêî = CisFun.Val2(H1lb_MaxKai)
        End With
        On Error GoTo DBPut_Err
        If AddFlg Then
            Call HCSInsert
        Else
            gSL_Select = "UPDATE î≠íçÉTÉCÉNÉãê›íËÉ}ÉXÉ^ SET"
            gSL_Select = gSL_Select & " íãñÈãÊï™ = " & HCS.íãñÈãÊï™ & ","
            gSL_Select = gSL_Select & " âÒé˚éûä‘ = '" & HCS.âÒé˚éûä‘ & "',"
            gSL_Select = gSL_Select & " î≠íçéûä‘ = '" & HCS.î≠íçéûä‘ & "',"
            gSL_Select = gSL_Select & " à¯éÊéûä‘ = '" & HCS.à¯éÊéûä‘ & "',"
            gSL_Select = gSL_Select & " îıçl = '" & HCS.îıçl & "',"
            gSL_Select = gSL_Select & " ç≈ëÂâÒêî = " & HCS.ç≈ëÂâÒêî & ","
            gSL_Select = gSL_Select & " çXêVì˙ = GETDATE(),"
            gSL_Select = gSL_Select & " çXêVí[ññ = HOST_NAME()"
            gSL_Select = gSL_Select & " WHERE âÒêî = " & gInt
            CisDB.SQL = gSL_Select
            CisDB.DBExec
        End If
    Next
    CisDB.DBTran (TransCommit) ' ƒ◊›ªﬁ∏ºÆ› èIóπ
    
'-----------------------------------'
'<< î≠íçÉTÉCÉNÉãÉ}ÉXÉ^ çXêVÅ@>>
'-----------------------------------'
    For gCnt1 = 1 To CisFun.Val2(iT1_Kensu)
        If RTrim(vsGrid.TextMatrix(gCnt1, 25)) <> "" And (RTrim(vsGrid.TextMatrix(gCnt1, 26)) <> "" Or HCSCHGFlg) Then
            Erase mBinTime: Erase mPrtDay: ReDim mBinTime(mMaxKaisu): Erase mBinKai: ReDim mBinKai(mMaxKaisu): ReDim mPrtDay(mMaxKaisu)
            For gCnt2 = 1 To mMaxKaisu
                If RTrim(vsGrid.TextMatrix(gCnt1, gCnt2 + 4)) <> "" Then
                    mBinTime(CisFun.Val2(vsGrid.TextMatrix(gCnt1, gCnt2 + 26))) = mHCTime(gCnt2)
                    mBinKai(CisFun.Val2(vsGrid.TextMatrix(gCnt1, gCnt2 + 26))) = gCnt2
                    mPrtDay(CisFun.Val2(vsGrid.TextMatrix(gCnt1, gCnt2 + 26))) = CisFun.Val2(vsGrid.TextMatrix(gCnt1, gCnt2 + 42))
                End If
            Next gCnt2
            BinMax = CisFun.Val2(Mid(vsGrid.TextMatrix(gCnt1, 24), 2, 2))
            BinDay = CisFun.Val2(Mid(vsGrid.TextMatrix(gCnt1, 24), 1, 1))
            mTorcd = RTrim(vsGrid.TextMatrix(gCnt1, 23))
            mUkeir = RTrim(vsGrid.TextMatrix(gCnt1, 2))
            mCycle = RTrim(vsGrid.TextMatrix(gCnt1, 24))
            Call FuriBin_Set
            
            gSL_Select = "DELETE FROM î≠íçÉTÉCÉNÉãÉ}ÉXÉ^"
            gSL_Select = gSL_Select & " WHERE édì¸êÊ = '" & RTrim(mTorcd) & "'"
            gSL_Select = gSL_Select & "   AND ÉTÉCÉNÉã = '" & RTrim(mCycle) & "'"
            gSL_Select = gSL_Select & "   AND éÛì¸ = '" & RTrim(mUkeir) & "'"
            CisDB.SQL = gSL_Select
            CisDB.DBExec
            
            For gIndex = 1 To BinMax + 1
                If RTrim(mSTTime(gIndex)) <> "" Then
                'ÉTÉCÉNÉã 1-99-99 , ÉTÉCÉNÉã 2-01-99 ÇÃñ≥Ç¢éûÅ@î≠íçÉTÉCÉNÉãÉ}ÉXÉ^
                    With HCM
                        .édì¸êÊ = RTrim(mTorcd)
                        .ÉTÉCÉNÉã = RTrim(mCycle)
                        .éÛì¸ = RTrim(mUkeir)
                        .êUèoï÷ = Format(mFurBin(gIndex), "00")
                        .çséØï  = mEdaNo(gIndex)
        
        
                        .íãñÈãÊï™ = mHYKbn(mBinKai(mFurBin(gIndex)))
                        .êUèoéûä‘ = mHCTime(mBinKai(mFurBin(gIndex)))
                        .î≠çséûä‘ = mHCTime(mBinKai(mFurBin(gIndex)))
                        .âÒé˚éûä‘ = mRETime(mBinKai(mFurBin(gIndex)))
                        .ì¸â◊éûä‘ = mHKTime(mBinKai(mFurBin(gIndex)))
                        .êUèoäJén = RTrim(mSTTime(gIndex))
                        .êUèoèIóπ = RTrim(mEDTime(gIndex))
                        .êUèoâ¡éZì˙ = mFurDay(gIndex)
                        .î≠íçâ¡éZì˙ = mHCDay(gIndex)
                        .î≠íçï÷ = Format(mHCBin(gIndex), "00")
                        .î≠çsâ¡éZì˙ = mPrtDay(mFurBin(gIndex))
        
                        .ç≈èIî≠íçì˙ = RTrim(vsGrid.TextMatrix(gCnt1, 21))
                        .ç≈èIî≠íçï÷ = Format(CisFun.Val2(vsGrid.TextMatrix(gCnt1, 22)), "00")
                        .î≠íçèàóùâÒêî = mBinKai(mFurBin(gIndex))
                        .ç≈èIï÷ = Format(CisFun.Val2(vsGrid.TextMatrix(gCnt1, 25)), "00")
        
                    End With
        '
                    HCM.çÏê¨é“ = gTanto
                ' ( í«â¡èàóù )
                    Call HCMInsert
                End If
            Next gIndex
        'ÉVÅ[ÉPÉìÉXï÷çÃî‘(íãñÈï÷èá>ñÈãŒÇÕíããŒë±Ç´ÇÃï÷Ç©ÇÁÅj
            Dim wFurBin()     As Byte
            Dim wEdaNo()      As Byte
            Dim wHYKbn()      As Byte
            gSL_Select = "SELECT * FROM î≠íçÉTÉCÉNÉãÉ}ÉXÉ^"
            gSL_Select = gSL_Select & " WHERE édì¸êÊ = '" & RTrim(mTorcd) & "'"
            gSL_Select = gSL_Select & "   And ÉTÉCÉNÉã = '" & RTrim(mCycle) & "'"
            gSL_Select = gSL_Select & "   And éÛì¸ = '" & RTrim(mUkeir) & "'"
            gSL_Select = gSL_Select & " ORDER BY êUèoï÷,çséØï "
            If HCMRead(gSL_Select, 0) Then
                ReDim wFurBin(CisDB.RecordCount)
                ReDim wEdaNo(CisDB.RecordCount)
                ReDim wHYKbn(CisDB.RecordCount)
                gIndex = 0
                Do Until Not HCM_RDSTS
                    gIndex = gIndex + 1
                    wFurBin(gIndex) = CisFun.Val2(HCM.êUèoï÷)
                    wEdaNo(gIndex) = HCM.çséØï 
                    wHYKbn(gIndex) = HCM.íãñÈãÊï™
                    Call HCMReadNext
                Loop
            End If
            Call HCMClose
            gSL_Select = "SELECT COUNT(*) åèêî FROM î≠íçÉTÉCÉNÉãÉ}ÉXÉ^"
            gSL_Select = gSL_Select & " WHERE édì¸êÊ = '" & RTrim(mTorcd) & "'"
            gSL_Select = gSL_Select & "   And ÉTÉCÉNÉã = '" & RTrim(mCycle) & "'"
            gSL_Select = gSL_Select & "   And éÛì¸ = '" & RTrim(mUkeir) & "'"
            gSL_Select = gSL_Select & "   And íãñÈãÊï™ = 0"
            If HCMRead(gSL_Select, 1) Then
            End If
            If HCM.åèêî <> 0 Then
                gCnt2 = 0
                Do Until wHYKbn(1) = 0
                    wFurBin(0) = wFurBin(1)
                    wEdaNo(0) = wEdaNo(1)
                    wHYKbn(0) = wHYKbn(1)
                    For gCnt2 = 1 To UBound(wFurBin) - 1
                        wFurBin(gCnt2) = wFurBin(gCnt2 + 1)
                        wEdaNo(gCnt2) = wEdaNo(gCnt2 + 1)
                        wHYKbn(gCnt2) = wHYKbn(gCnt2 + 1)
                    Next
                    wFurBin(UBound(wFurBin)) = wFurBin(0)
                    wEdaNo(UBound(wFurBin)) = wEdaNo(0)
                    wHYKbn(UBound(wFurBin)) = wHYKbn(0)
                Loop
            End If
            gCnt2 = 0
            For gIndex = 1 To UBound(wFurBin)
                If wEdaNo(gIndex) = 0 Then
                    gCnt2 = gCnt2 + 1
                End If
                With CisDB
                    On Error GoTo DBPut_Err
                    .SQL = "UPDATE î≠íçÉTÉCÉNÉãÉ}ÉXÉ^ SET"
                    .SQL = .SQL & " ÉVÅ[ÉPÉìÉXï÷ = '" & Format(gCnt2, "00") & "'"
                    .SQL = .SQL & " WHERE édì¸êÊ = '" & RTrim(mTorcd) & "'"
                    .SQL = .SQL & "   And ÉTÉCÉNÉã = '" & RTrim(mCycle) & "'"
                    .SQL = .SQL & "   And éÛì¸ = '" & RTrim(mUkeir) & "'"
                    .SQL = .SQL & "   And êUèoï÷ = '" & Format(wFurBin(gIndex), "00") & "'"
                    .SQL = .SQL & "   And çséØï  = " & wEdaNo(gIndex) & ""
                    .DBExec
            
                    On Error GoTo 0
                End With
            Next
        'êUèoâ¡éZì˙ÇÃí˘ê≥(ÇPï÷Ç™ñÈãŒÇ©ÇÁénÇ‹ÇËÇQï÷à»ç~íããŒÇ™óLÇÈèÍçáÅj
            gSL_Select = "SELECT * FROM î≠íçÉTÉCÉNÉãÉ}ÉXÉ^"
            gSL_Select = gSL_Select & " WHERE édì¸êÊ = '" & RTrim(mTorcd) & "'"
            gSL_Select = gSL_Select & "   And ÉTÉCÉNÉã = '" & RTrim(mCycle) & "'"
            gSL_Select = gSL_Select & "   And éÛì¸ = '" & RTrim(mUkeir) & "'"
            gSL_Select = gSL_Select & " ORDER BY êUèoï÷,çséØï "
            If HCMRead(gSL_Select, 0) Then
                ReDim wFurBin(CisDB.RecordCount)
                ReDim wEdaNo(CisDB.RecordCount)
                ReDim wHYKbn(CisDB.RecordCount)
                gIndex = 0: wYDJun = 0
                Do Until Not HCM_RDSTS
                    gIndex = gIndex + 1
                    wFurBin(gIndex) = CisFun.Val2(HCM.êUèoï÷)
                    wEdaNo(gIndex) = HCM.çséØï 
                    wHYKbn(gIndex) = HCM.íãñÈãÊï™
                    If wYDJun = 0 And wHYKbn(gIndex) = 1 Then wYDJun = 1
                    If wYDJun = 1 And wHYKbn(gIndex) = 0 Then wYDJun = 2
                    Call HCMReadNext
                Loop
            End If
            Call HCMClose
            If wYDJun = 2 Then
                For gIndex = 1 To UBound(wFurBin)
                    If wHYKbn(gIndex) = 0 Then Exit For
                    With CisDB
                        On Error GoTo DBPut_Err
                        .SQL = "UPDATE î≠íçÉTÉCÉNÉãÉ}ÉXÉ^ SET"
                        .SQL = .SQL & " êUèoâ¡éZì˙ = êUèoâ¡éZì˙ - 1"
                        .SQL = .SQL & " WHERE édì¸êÊ = '" & RTrim(mTorcd) & "'"
                        .SQL = .SQL & "   And ÉTÉCÉNÉã = '" & RTrim(mCycle) & "'"
                        .SQL = .SQL & "   And éÛì¸ = '" & RTrim(mUkeir) & "'"
                        .SQL = .SQL & "   And êUèoï÷ = '" & Format(wFurBin(gIndex), "00") & "'"
                        .SQL = .SQL & "   And çséØï  = " & wEdaNo(gIndex) & ""
                        .DBExec
                
                        On Error GoTo 0
                    End With
                Next gIndex
            End If
        Else
            If RTrim(vsGrid.TextMatrix(gCnt1, 25)) = "" And RTrim(vsGrid.TextMatrix(gCnt1, 26)) <> "" Then
                'ê›íËÇ™Ç»Ç≠Ç»Ç¡ÇΩèÍçáÇÕçÌèú
                mTorcd = RTrim(vsGrid.TextMatrix(gCnt1, 23))
                mUkeir = RTrim(vsGrid.TextMatrix(gCnt1, 2))
                mCycle = RTrim(vsGrid.TextMatrix(gCnt1, 24))
                With CisDB
                    On Error GoTo DBPut_Err
                    .SQL = "DELETE FROM î≠íçÉTÉCÉNÉãÉ}ÉXÉ^ "
                    .SQL = .SQL & " WHERE édì¸êÊ = '" & RTrim(mTorcd) & "'"
                    .SQL = .SQL & "   And ÉTÉCÉNÉã = '" & RTrim(mCycle) & "'"
                    .SQL = .SQL & "   And éÛì¸ = '" & RTrim(mUkeir) & "'"
                    .DBExec
            
                    On Error GoTo 0
                End With
            End If
        End If
    Next gCnt1
    
    With CisDB
        On Error GoTo DBPut_Err
        .SQL = "DELETE FROM î≠íçÉTÉCÉNÉãÉ}ÉXÉ^"
        .SQL = .SQL & " FROM î≠íçÉTÉCÉNÉãÉ}ÉXÉ^ HC"
        .SQL = .SQL & " LEFT OUTER JOIN (SELECT éÊà¯êÊ,éÛì¸,ÉTÉCÉNÉã FROM ïiî‘éÊà¯êÊÉ}ÉXÉ^"
        .SQL = .SQL & "             WHERE ÉTÉCÉNÉã <> ''"
        .SQL = .SQL & "             GROUP BY éÊà¯êÊ,éÛì¸,ÉTÉCÉNÉã) HT"
        .SQL = .SQL & "     ON HT.éÊà¯êÊ = HC.édì¸êÊ"
        .SQL = .SQL & "     AND HT.éÛì¸ = HC.éÛì¸"
        .SQL = .SQL & "     AND HT.ÉTÉCÉNÉã = HC.ÉTÉCÉNÉã"
        .SQL = .SQL & " Where HT.éÊà¯êÊ Is Null"
        .DBExec

        On Error GoTo 0
    End With
    
    
    
    
    PNL_Wait.Visible = False
    DoEvents
    
    
    DBPut = True
    Exit Function
DBPut_Err:
    PNL_Wait.Visible = False
    DoEvents
    CisFun.ErrorBox
    End
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
        .Clear
'       Call .FixedSet(å©èoÇµà íuê›íË,óÒà íuê›íË,œ∞ºﬁèàóù,
'                      æŸåÖêî,óÒî‘çÜ,å©èoÇµì‡óe,ï™óﬁ¿≤Ãﬂ)
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 3, 0, "áÇ")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 13, 1, "édì¸êÊ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 2, 2, "éÛ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 8, 3, "ÉTÉCÉNÉã")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 16, 4, "ç≈èIî≠íçì˙-ï÷")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 5, 5, "1âÒ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 5, 6, "2âÒ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 5, 7, "3âÒ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 5, 8, "4âÒ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 5, 9, "5âÒ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 5, 10, "6âÒ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 5, 11, "7âÒ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 5, 12, "8âÒ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 5, 13, "9âÒ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 5, 14, "10âÒ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 5, 15, "11âÒ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 5, 16, "12âÒ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 5, 17, "13âÒ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 5, 18, "14âÒ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 5, 19, "15âÒ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 5, 20, "16âÒ")
'Å@∏ﬁÿØƒﬁëIëèàóùópÅ@ëﬁî¥ÿ±
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 0, 21, "Xç≈èIî≠íçì˙")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 0, 22, "Xç≈èIî≠íçï÷")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 0, 23, "Xédì¸êÊ")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 0, 24, "XÉTÉCÉNÉã")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 0, 25, "ç≈èIï÷")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 0, 26, "ïœçX")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 27, "1ï÷")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 28, "2ï÷")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 29, "3ï÷")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 30, "4ï÷")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 31, "5ï÷")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 32, "6ï÷")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 33, "7ï÷")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 34, "8ï÷")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 35, "9ï÷")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 36, "10ï÷")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 37, "11ï÷")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 38, "12ï÷")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 39, "13ï÷")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 40, "14ï÷")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 41, "15ï÷")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 5, 42, "16ï÷")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 43, "1ëOì˙")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 44, "2ëOì˙")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 45, "3ëOì˙")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 46, "4ëOì˙")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 47, "5ëOì˙")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 48, "6ëOì˙")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 49, "7ëOì˙")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 50, "8ëOì˙")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 51, "9ëOì˙")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 52, "10ëOì˙")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 53, "11ëOì˙")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 54, "12ëOì˙")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 55, "13ëOì˙")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 56, "14ëOì˙")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 57, "14ëOì˙")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 0, 58, "14ëOì˙")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 0, 59, "Select")
'  ï\é¶ ﬂ¿∞›æØƒ
'       Call .PatanSet( ﬂ¿∞›NO,ï\é¶ÉpÉ^Å[Éì)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20")
        .InitDisp
    End With
    
End Sub
'******************************************'
'*****       ÉOÉäÉbÉhì‡óeÉZÉbÉg        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Long
    Dim Key_Torcd   As String
    Dim Key_Cycle   As String
    Dim Key_Ukeir   As String
' íäèoèàóù èâä˙æØƒ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' íäèoåèêî(çáåv)ëﬁî
' ∏ﬁÿØƒﬁçsçƒæØƒ & óÒà íuéÊìæ
    With vsGrid
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(CisFun.Val2(ID))
        Next ID
        If Ttl_Kensu + 1 > .Rows Then
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not HCM_RDSTS
        'Grid Row,Col
        If Key_Torcd <> HCM.édì¸êÊ Or Key_Cycle <> HCM.ÉTÉCÉNÉã Or Key_Ukeir <> HCM.éÛì¸ Then
            ID = ID + 1
            With vsGrid
                .TextMatrix(ID, GridCol(0)) = Format(ID, "##")
                .TextMatrix(ID, GridCol(1)) = RTrim(HCM.édì¸êÊñº)
                .TextMatrix(ID, GridCol(2)) = HCM.éÛì¸
                If HCM.ójì˙î≠íç = 0 Then
                    .TextMatrix(ID, GridCol(3)) = HCM.ï“èWÉTÉCÉNÉã
                Else
                    .TextMatrix(ID, GridCol(3)) = EditWeekDay
                End If
                If RTrim(HCM.ç≈èIî≠íçì˙) <> "" Then
                    gStr = Left(HCM.ç≈èIî≠íçì˙, 4) & "/" & Mid(HCM.ç≈èIî≠íçì˙, 5, 2) & "/" & Mid(HCM.ç≈èIî≠íçì˙, 7, 2)
                Else
                    gStr = ""
                End If
                If RTrim(HCM.ç≈èIî≠íçï÷) <> "" And RTrim(HCM.ç≈èIî≠íçï÷) <> "00" Then
                    gStr = gStr & "-" & RTrim(HCM.ç≈èIî≠íçï÷)
                End If
                .TextMatrix(ID, GridCol(4)) = gStr
                .TextMatrix(ID, GridCol(21)) = HCM.ç≈èIî≠íçì˙
                .TextMatrix(ID, GridCol(22)) = HCM.ç≈èIî≠íçï÷
                .TextMatrix(ID, GridCol(23)) = HCM.édì¸êÊ
                .TextMatrix(ID, GridCol(24)) = HCM.ÉTÉCÉNÉã
                .TextMatrix(ID, GridCol(25)) = ""
                .TextMatrix(ID, GridCol(26)) = ""
            End With
        End If
        If HCM.î≠íçèàóùâÒêî <> 0 Then
            vsGrid.TextMatrix(ID, HCM.î≠íçèàóùâÒêî + 26) = CisFun.Val2(HCM.êUèoï÷)
            vsGrid.TextMatrix(ID, HCM.î≠íçèàóùâÒêî + 42) = CisFun.Val2(HCM.î≠çsâ¡éZì˙)
            If HCM.î≠çsâ¡éZì˙ = 0 Then
                vsGrid.TextMatrix(ID, HCM.î≠íçèàóùâÒêî + 4) = HCM.êUèoï÷
            Else
                vsGrid.TextMatrix(ID, HCM.î≠íçèàóùâÒêî + 4) = HCM.êUèoï÷ & "*"
            End If
        End If
        
        Key_Torcd = HCM.édì¸êÊ
        Key_Cycle = HCM.ÉTÉCÉNÉã
        Key_Ukeir = HCM.éÛì¸
        
        Call HCMReadNext
    Loop

    With vsGrid
         .Rows = ID + 1
        iT1_Kensu = Format(ID, "#,###")
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'==========================================================='
'   ójì˙î≠íçÅ@ï\é¶
'==========================================================='
Private Function EditWeekDay() As String
    EditWeekDay = ""
    gSL_Select = "SELECT * FROM éÊà¯êÊéÛì¸É}ÉXÉ^"
    gSL_Select = gSL_Select & " WHERE éÊà¯êÊCD = '" & RTrim(HCM.édì¸êÊ) & "'"
    gSL_Select = gSL_Select & "   AND éÛì¸ = '" & RTrim(HCM.éÛì¸) & "'"
    gSL_Select = gSL_Select & "   AND éÊà¯êÊãÊï™ = 1"
    If Not TUMRead(gSL_Select, 1, 1) Then Exit Function
    If TUM.ójì˙1 <> 0 Then EditWeekDay = EditWeekDay & "åé"
    If TUM.ójì˙2 <> 0 Then EditWeekDay = EditWeekDay & "âŒ"
    If TUM.ójì˙3 <> 0 Then EditWeekDay = EditWeekDay & "êÖ"
    If TUM.ójì˙4 <> 0 Then EditWeekDay = EditWeekDay & "ñÿ"
    If TUM.ójì˙5 <> 0 Then EditWeekDay = EditWeekDay & "ã‡"

End Function
'==========================================================='
'   ì¸óÕÉeÉLÉXÉgï\é¶ÅEîÒï\é¶
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
    If SyoriKB = "MNT" Then
        Input_Row.Visible = EnabledFlg
        B_TabDummy.Enabled = EnabledFlg
        B_TabDummy2.Enabled = EnabledFlg
    Else
        Input_Row.Visible = False
        B_TabDummy.Enabled = False
        B_TabDummy2.Enabled = False
    End If
End Sub
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub

'ÉeÉLÉXÉgÇ©ÇÁÉOÉäÉbÉhÇ…ëﬁî
    Call TextToVsGrid(Val(Input_Row.Tag))
    gInt = vsGrid.Row
    'Å™
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    'Å´
    If KeyCode = vbKeyDown Then
        If gInt + 1 > vsGrid.Rows - 1 Then GoTo Grid_RowChange_End
        If vsGrid.TextMatrix(gInt + 1, 0) = "" Then GoTo Grid_RowChange_End
        gInt = gInt + 1
    End If
    'PageUp
    If KeyCode = vbKeyPageUp Then
        If gInt - CisVsGrid.DispMax < 1 Then
            gInt = 1
        Else
            gInt = gInt - CisVsGrid.DispMax
        End If
    End If
    'PageDown
    If KeyCode = vbKeyPageDown Then
        If gInt + CisVsGrid.DispMax > vsGrid.Rows - 1 Then
            gInt = vsGrid.Rows - 1
        Else
            gInt = gInt + CisVsGrid.DispMax
        End If
    End If
    vsGrid.Row = gInt
'ÉOÉäÉbÉhÇ©ÇÁÉeÉLÉXÉgÇ…ï\é¶
    Call VsGridToText(vsGrid.Row)
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   ì¸óÕÉeÉLÉXÉgÇ©ÇÁÉOÉäÉbÉhÇ…à⁄Ç∑
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    With vsGrid
        If RTrim(iB1_EHYmd) <> "" Then
            .TextMatrix(tRow, 21) = RTrim(iB1_EHYmd)
            If iB1_EHYmd.cChkResult Then
                With CisFun
                    .DateE_BefVal = RTrim(iB1_EHYmd)
                    .DateEdit Slash
                    vsGrid.TextMatrix(tRow, 4) = .DateE_AfVal
                End With
            Else
                vsGrid.TextMatrix(tRow, 4) = RTrim(iB1_EHYmd)
            End If
        Else
            .TextMatrix(tRow, 21) = ""
            .TextMatrix(tRow, 4) = ""
        End If
        If CisFun.Val2(iB1_EHBin) <> 0 Then
            .TextMatrix(tRow, 22) = RTrim(iB1_EHBin)
            .TextMatrix(tRow, 4) = .TextMatrix(tRow, 4) & "-" & Format(CisFun.Val2(iB1_EHBin), "00")
        Else
            .TextMatrix(tRow, 22) = ""
        End If
    
        For gInt = 1 To mMaxKaisu
            If CisFun.Val2(iB1_Bin(gInt)) <> 0 Then
                .TextMatrix(tRow, gInt + 26) = iB1_Bin(gInt)
            Else
                .TextMatrix(tRow, gInt + 26) = ""
            End If
            .TextMatrix(tRow, gInt + 42) = ChkB1_ZHKbn(gInt).Value
            If CisFun.Val2(iB1_Bin(gInt)) <> 0 Then
                If ChkB1_ZHKbn(gInt).Value = 0 Then
                    .TextMatrix(tRow, gInt + 4) = iB1_Bin(gInt)
                Else
                    .TextMatrix(tRow, gInt + 4) = iB1_Bin(gInt) & "*"
                End If
            Else
                .TextMatrix(tRow, gInt + 4) = ""
            End If
        Next
        
    End With
End Sub
'==========================================================='
'   ÉOÉäÉbÉhÇ©ÇÁì¸óÕÉeÉLÉXÉgÇ…à⁄Ç∑ÅAì¸óÕÉeÉLÉXÉgÇÃà⁄ìÆ
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With vsGrid
        Input_Row.Tag = tRow
        Input_Row.Move Input_Row.Left, vsGrid.CellTop, Input_Row.Width, vsGrid.CellHeight
        'ç≈èIî≠íçì˙ÅEï÷
        iB1_EHYmd = .TextMatrix(tRow, 21)
        iB1_EHBin = .TextMatrix(tRow, 22)
       'êUèoéûä‘
        For gInt = 1 To mMaxKaisu
            iB1_Bin(gInt) = .TextMatrix(tRow, gInt + 26)
            ChkB1_ZHKbn(gInt).Value = CisFun.Val2(.TextMatrix(tRow, gInt + 42))
        Next
    End With
End Sub
Private Sub VSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
    If SyoriKB <> "MNT" Then Exit Sub
    If ProcHB <> "B1" Then Exit Sub
    Call TextToVsGrid(Val(Input_Row.Tag))
    Call VsGridToText(vsGrid.Row)
End Sub
Private Sub VSGrid_Click()
    If SyoriKB <> "MNT" Then Exit Sub
    If ProcHB <> "B1" Then Exit Sub
    Call TextToVsGrid(Val(Input_Row.Tag))
    If vsGrid.TextMatrix(vsGrid.Row, 0) = "" Then Exit Sub
    Call GridTextEnabled(True)
    Call VsGridToText(vsGrid.Row)
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub
Private Sub vsGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SyoriKB <> "KAN" Then Exit Sub
    If ProcHB <> "B1" Then Exit Sub
    cMouseRow = vsGrid.MouseRow
    cMouseCol = vsGrid.MouseCol
'-------------------------------------- 2006/03/29 Insert
    If cMouseRow = 0 Then Exit Sub
    If cMouseCol = 1 And Button = 2 Then
        With vsGrid
             If cMouseRow = 0 Then Exit Sub
             .Row = .MouseRow
             .Col = .MouseCol
              DispTorcd = .TextMatrix(cMouseRow, CisVsGrid.FixedGet(23))
              DispTorcd = DispTorcd & ":" & .TextMatrix(cMouseRow, CisVsGrid.FixedGet(1))
        End With
        
        DispTorcd.Move vsGrid.Left + X, Y - DispTorcd.Height
        DispTorcd.Visible = True
        Exit Sub
    End If
'-------------------------------------- 2006/03/29 Insert End
    
    If cMouseCol < 4 Or cMouseCol > 20 Then Exit Sub
    
    If cMouseCol <> 4 Then
        If Button = 1 Then
            Call ClickBinDisp
        Else
            vsGrid.TextMatrix(cMouseRow, cMouseCol) = ""
            vsGrid.TextMatrix(cMouseRow, cMouseCol + 22) = ""
            vsGrid.TextMatrix(cMouseRow, cMouseCol + 38) = "0"
        End If
    Else
        Call ClickHYmdDisp(Button)
    End If
    vsGrid.TextMatrix(cMouseRow, 26) = "M"
End Sub

Private Sub vsGrid_Scroll()
    'ÉXÉNÉçÅ[ÉãéûÇÕÉeÉLÉXÉgÉ{ÉbÉNÉXÇîÒï\é¶Ç…ÇµÇ‹Ç∑ÅB
'    Call GridTextEnabled(False)
End Sub
'-------------------------------------- 2006/03/29 Insert
Private Sub VSGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DispTorcd.Visible = False
End Sub
'-------------------------------------- 2006/03/29 Insert End
'-----------------------------------------------------------------------'
'   ç≈ëÂï÷ÇãÅÇﬂÅAç≈ëÂï÷+ÇPÇÉOÉäÉbÉhÇ…ï\é¶Ç∑ÇÈ
'-----------------------------------------------------------------------'
Private Sub ClickBinDisp()
    gCnt1 = 0
    BinMax = CisFun.Val2(Mid(vsGrid.TextMatrix(cMouseRow, 24), 2, 2))
    If vsGrid.TextMatrix(cMouseRow, cMouseCol) = "" Then
        For gInt = 27 To MaxKaisu + 27
            If CisFun.Val2(vsGrid.TextMatrix(cMouseRow, gInt)) <> 0 Then
                If gCnt1 < CisFun.Val2(vsGrid.TextMatrix(cMouseRow, gInt)) Then
                    gCnt1 = CisFun.Val2(vsGrid.TextMatrix(cMouseRow, gInt))
                End If
            End If
        Next
        If gCnt1 < BinMax Then
            vsGrid.TextMatrix(cMouseRow, cMouseCol) = Format(gCnt1 + 1, "00")
            vsGrid.TextMatrix(cMouseRow, cMouseCol + 22) = Format(gCnt1 + 1, "00")
        End If
    Else
        If CisFun.Val2(vsGrid.TextMatrix(cMouseRow, cMouseCol + 38)) = 0 Then
            vsGrid.TextMatrix(cMouseRow, cMouseCol + 38) = 1
            vsGrid.TextMatrix(cMouseRow, cMouseCol) = vsGrid.TextMatrix(cMouseRow, cMouseCol) & "*"
        Else
            vsGrid.TextMatrix(cMouseRow, cMouseCol + 38) = 0
            vsGrid.TextMatrix(cMouseRow, cMouseCol) = Replace(vsGrid.TextMatrix(cMouseRow, cMouseCol), "*", "")
        End If
    End If
End Sub
'-----------------------------------------------------------------------'
'   ç≈èIî≠íçì˙ÅEï÷Çï\é¶Ç∑ÇÈ
'       ãÛîíéûÅ@ÅcÅ@ÉVÉXÉeÉÄì˙ït
'       ç∂ÉNÉäÉbÉNÅ@ÅcÅ@1ï÷â¡éZ
'       âEÉNÉäÉbÉNÅ@ÅcÅ@1ï÷å∏éZ
'-----------------------------------------------------------------------'
Private Sub ClickHYmdDisp(wButton As Integer)
    Static wBin As Integer
    Static wHYmd As Date

    If RTrim(vsGrid.TextMatrix(cMouseRow, 21)) = "" Then
        vsGrid.TextMatrix(cMouseRow, cMouseCol) = Format(Now(), "yyyy/mm/dd") & "-01"
        vsGrid.TextMatrix(cMouseRow, 21) = Format(Now(), "yyyymmdd")
        vsGrid.TextMatrix(cMouseRow, 22) = "01"
        Exit Sub
    End If
    wHYmd = CDate(Left(vsGrid.TextMatrix(cMouseRow, 21), 4) & "/" & Mid(vsGrid.TextMatrix(cMouseRow, 21), 5, 2) & "/" & Mid(vsGrid.TextMatrix(cMouseRow, 21), 7, 2))
    wBin = CisFun.Val2(vsGrid.TextMatrix(cMouseRow, 22))
    gCnt1 = 0
    BinMax = CisFun.Val2(Mid(vsGrid.TextMatrix(cMouseRow, 24), 2, 2))
    If wButton = 1 Then
        If wBin + 1 > BinMax Then
            wHYmd = DateAdd("d", 1, wHYmd)
            wBin = 1
        Else
            wBin = wBin + 1
        End If
    Else
        If wBin - 1 < 1 Then
            wHYmd = DateAdd("d", -1, wHYmd)
            wBin = BinMax
        Else
            wBin = wBin - 1
        End If
    End If

    vsGrid.TextMatrix(cMouseRow, cMouseCol) = Format(wHYmd, "yyyy/mm/dd") & "-" & Format(wBin, "00")
    vsGrid.TextMatrix(cMouseRow, 21) = Format(wHYmd, "yyyymmdd")
    vsGrid.TextMatrix(cMouseRow, 22) = Format(wBin, "00")

End Sub

