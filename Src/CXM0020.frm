VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXM0020 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "éÊà¯êÊÉ}ÉXÉ^ÉÅÉìÉeÉiÉìÉX"
   ClientHeight    =   11190
   ClientLeft      =   2925
   ClientTop       =   930
   ClientWidth     =   15240
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z µ∞¿ﬁ∞
   ScaleHeight     =   11190
   ScaleWidth      =   15240
   WindowState     =   2  'ç≈ëÂâª
   Begin Cis3D_v60.CIS3D Back_Kotei 
      Height          =   7770
      Left            =   1620
      Top             =   3120
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   13705
      BackColor       =   16711680
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
      cBoderWidth     =   3
      Begin Cis3D_v60.CIS3D CIS3D24 
         Height          =   7515
         Left            =   120
         Top             =   120
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   13256
         BackColor       =   16711680
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
         Begin Cis3D_v60.CIS3D CIS3D35 
            Height          =   450
            Left            =   30
            Top             =   30
            Width           =   10920
            _ExtentX        =   19262
            _ExtentY        =   794
            BackColor       =   16761024
            Caption         =   "Åy çH íˆ èÓ ïÒ ìo ò^ Åz"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   18
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cBoderWidth     =   2
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D25 
            Height          =   3765
            Left            =   30
            Top             =   2895
            Width           =   10920
            _ExtentX        =   19262
            _ExtentY        =   6641
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
            Begin vsFlexLib.vsFlexArray VsGrid2 
               Height          =   3660
               Left            =   60
               TabIndex        =   97
               TabStop         =   0   'False
               Top             =   45
               Width           =   10785
               _Version        =   196608
               _ExtentX        =   19024
               _ExtentY        =   6456
               _StockProps     =   228
               Appearance      =   1
               ConvInfo        =   1418783674
               FormatString    =   "No|ã@äB|ã@äBñºÅ@Å@|ã@äBáÇ|ã@äBáÇñº  |â“ì≠éûä‘|ë‰êî|ë„ã@äB|ã@äBñºÅ@Å@|ë„ã@äBáÇ|ã@äBáÇñº  "
               Rows            =   4
               Cols            =   12
               BackColor       =   12648384
               BackColorBkg    =   12632256
               FocusRect       =   0
               HighLight       =   0
               RowHeightMin    =   330
               SelectionMode   =   1
               BackColorAlternate=   12648384
               Begin CisText_V60.CisText iB1_Kikai 
                  Height          =   360
                  Left            =   330
                  TabIndex        =   98
                  Top             =   330
                  Width           =   570
                  _ExtentX        =   1005
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
               Begin CisText_V60.CisText iB1_KadoTime 
                  Height          =   345
                  Left            =   4290
                  TabIndex        =   100
                  Top             =   330
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   609
                  cFormat         =   "##0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "##0.00"
                  MaxLength       =   6
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "##0.00"
                  cILength        =   3
               End
               Begin CisText_V60.CisText iB1_DaiKikai 
                  Height          =   360
                  Left            =   5910
                  TabIndex        =   102
                  Top             =   330
                  Width           =   825
                  _ExtentX        =   1455
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
               Begin Cis3D_v60.CIS3D B1lb_KikaiNM 
                  Height          =   300
                  Left            =   900
                  Top             =   330
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   529
                  BackColor       =   16777152
                  Caption         =   "300t"
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
                  cFont3D         =   2
                  cAlingnment     =   1
                  cBoderStyle     =   1
                  cPositionX      =   30
               End
               Begin CisText_V60.CisText iB1_Daisu 
                  Height          =   360
                  Left            =   5325
                  TabIndex        =   101
                  Top             =   330
                  Width           =   600
                  _ExtentX        =   1058
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
               Begin CisText_V60.CisText iB1_KikaiNo 
                  Height          =   360
                  Left            =   2190
                  TabIndex        =   99
                  Top             =   330
                  Width           =   810
                  _ExtentX        =   1429
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
               Begin Cis3D_v60.CIS3D B1lb_KikaiNoNM 
                  Height          =   300
                  Left            =   3000
                  Top             =   330
                  Width           =   1290
                  _ExtentX        =   2275
                  _ExtentY        =   529
                  BackColor       =   16777152
                  Caption         =   "P56"
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
                  cFont3D         =   2
                  cAlingnment     =   1
                  cBoderStyle     =   1
                  cPositionX      =   30
               End
               Begin Cis3D_v60.CIS3D B1lb_DaiKikaiNM 
                  Height          =   315
                  Left            =   6720
                  Top             =   330
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   556
                  BackColor       =   16777152
                  Caption         =   "300t"
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
                  cFont3D         =   2
                  cAlingnment     =   1
                  cBoderStyle     =   1
                  cPositionX      =   30
               End
               Begin CisText_V60.CisText iB1_DaiKikaiNo 
                  Height          =   360
                  Left            =   8010
                  TabIndex        =   103
                  Top             =   330
                  Width           =   1050
                  _ExtentX        =   1852
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
                  cAutoEject      =   0   'False
               End
               Begin Cis3D_v60.CIS3D B1lb_DaiKikaiNoNM 
                  Height          =   300
                  Left            =   9060
                  Top             =   330
                  Width           =   1290
                  _ExtentX        =   2275
                  _ExtentY        =   529
                  BackColor       =   16777152
                  Caption         =   "P55"
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
                  cFont3D         =   2
                  cAlingnment     =   1
                  cBoderStyle     =   1
                  cPositionX      =   30
               End
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hinbn 
            Height          =   810
            Left            =   30
            Top             =   6675
            Width           =   10920
            _ExtentX        =   19262
            _ExtentY        =   1429
            ForeColor       =   16711680
            Caption         =   "  ì˙ïÒèoóÕ"
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
            cPositionX      =   90
            cPositionY      =   30
            Begin VB.CheckBox B1Chk_PSChk 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ÉvÉåÉX"
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
               Left            =   4005
               TabIndex        =   106
               Top             =   420
               Width           =   1185
            End
            Begin VB.CheckBox B1Chk_Nipo 
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
               Height          =   255
               Left            =   300
               TabIndex        =   104
               Top             =   420
               Width           =   1185
            End
            Begin VB.CheckBox B1Chk_Fuka 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Ç∑ÇÈ"
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
               Left            =   1980
               TabIndex        =   105
               Top             =   405
               Width           =   1185
            End
            Begin Cis3D_v60.CIS3D CIS3D26 
               Height          =   285
               Left            =   2040
               Top             =   45
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               ForeColor       =   16711680
               Caption         =   "ïââ◊ä«óù"
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
            Begin Cis3D_v60.CIS3D CIS3D27 
               Height          =   285
               Left            =   3495
               Top             =   60
               Width           =   1950
               _ExtentX        =   3440
               _ExtentY        =   503
               ForeColor       =   16711680
               Caption         =   "ÉvÉåÉX/ÉXÉ|ÉbÉg"
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
         End
         Begin Cis3D_v60.CIS3D CIS3D28 
            Height          =   1755
            Left            =   30
            Top             =   1125
            Width           =   10920
            _ExtentX        =   19262
            _ExtentY        =   3096
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
            cAlingnment     =   6
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D CIS3D29 
               Height          =   435
               Left            =   75
               Top             =   375
               Width           =   5130
               _ExtentX        =   9049
               _ExtentY        =   767
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
               Begin Cis3D_v60.CIS3D CIS3D30 
                  Height          =   375
                  Left            =   30
                  Top             =   30
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   661
                  ForeColor       =   16711680
                  Caption         =   "ê≥é–àı"
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
               End
               Begin CisText_V60.CisText iB1_STime 
                  Height          =   375
                  Left            =   2580
                  TabIndex        =   89
                  Top             =   30
                  Width           =   810
                  _ExtentX        =   1429
                  _ExtentY        =   661
                  cFormat         =   "##0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "##0.00"
                  MaxLength       =   6
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "##0.00"
                  cILength        =   3
               End
               Begin CisText_V60.CisText iB1_SGenka 
                  Height          =   375
                  Left            =   3900
                  TabIndex        =   90
                  Top             =   30
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   661
                  cFormat         =   "####0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "####0.00"
                  MaxLength       =   8
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "####0.00"
                  cILength        =   5
               End
               Begin CisText_V60.CisText iB1_SHito 
                  Height          =   375
                  Left            =   1305
                  TabIndex        =   88
                  Top             =   30
                  Width           =   735
                  _ExtentX        =   1296
                  _ExtentY        =   661
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
            Begin Cis3D_v60.CIS3D CIS3D31 
               Height          =   315
               Left            =   1350
               Top             =   30
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               ForeColor       =   16711680
               Caption         =   "êlàı"
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
               cAlingnment     =   7
               cBoderStyle     =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D32 
               Height          =   315
               Left            =   2415
               Top             =   30
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               ForeColor       =   16711680
               Caption         =   "â“ì≠éûä‘"
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
               cAlingnment     =   7
               cBoderStyle     =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D33 
               Height          =   435
               Left            =   75
               Top             =   810
               Width           =   5130
               _ExtentX        =   9049
               _ExtentY        =   767
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
               Begin Cis3D_v60.CIS3D CIS3D34 
                  Height          =   375
                  Left            =   30
                  Top             =   30
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   661
                  ForeColor       =   16711680
                  Caption         =   "ÉpÅ[Ég"
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
               End
               Begin CisText_V60.CisText iB1_PHito 
                  Height          =   375
                  Left            =   1305
                  TabIndex        =   91
                  Top             =   30
                  Width           =   735
                  _ExtentX        =   1296
                  _ExtentY        =   661
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
               Begin CisText_V60.CisText iB1_PTime 
                  Height          =   375
                  Left            =   2580
                  TabIndex        =   92
                  Top             =   30
                  Width           =   810
                  _ExtentX        =   1429
                  _ExtentY        =   661
                  cFormat         =   "##0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "##0.00"
                  MaxLength       =   6
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "##0.00"
                  cILength        =   3
               End
               Begin CisText_V60.CisText iB1_PGenka 
                  Height          =   375
                  Left            =   3900
                  TabIndex        =   93
                  Top             =   30
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   661
                  cFormat         =   "####0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "####0.00"
                  MaxLength       =   8
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "####0.00"
                  cILength        =   5
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D36 
               Height          =   435
               Left            =   75
               Top             =   1245
               Width           =   5130
               _ExtentX        =   9049
               _ExtentY        =   767
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
               Begin Cis3D_v60.CIS3D CIS3D37 
                  Height          =   375
                  Left            =   30
                  Top             =   30
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   661
                  ForeColor       =   16711680
                  Caption         =   "ÇªÇÃëº"
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
               End
               Begin CisText_V60.CisText iB1_THito 
                  Height          =   375
                  Left            =   1305
                  TabIndex        =   94
                  Top             =   30
                  Width           =   735
                  _ExtentX        =   1296
                  _ExtentY        =   661
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
               Begin CisText_V60.CisText iB1_TTime 
                  Height          =   375
                  Left            =   2580
                  TabIndex        =   95
                  Top             =   30
                  Width           =   810
                  _ExtentX        =   1429
                  _ExtentY        =   661
                  cFormat         =   "##0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "##0.00"
                  MaxLength       =   6
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "##0.00"
                  cILength        =   3
               End
               Begin CisText_V60.CisText iB1_TGenka 
                  Height          =   375
                  Left            =   3900
                  TabIndex        =   96
                  Top             =   30
                  Width           =   1080
                  _ExtentX        =   1905
                  _ExtentY        =   661
                  cFormat         =   "####0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "####0.00"
                  MaxLength       =   8
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "####0.00"
                  cILength        =   5
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D38 
               Height          =   315
               Left            =   3915
               Top             =   30
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               ForeColor       =   16711680
               Caption         =   "å¥âø/h"
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
               cAlingnment     =   7
               cBoderStyle     =   2
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   615
            Index           =   1
            Left            =   30
            Top             =   495
            Width           =   10920
            _ExtentX        =   19262
            _ExtentY        =   1085
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
            cAlingnment     =   6
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D PB_KoteiEND 
               Height          =   585
               Left            =   9540
               Top             =   15
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   1032
               ForeColor       =   16711680
               Caption         =   "ï¬Ç∂ÇÈ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cBoderWidth     =   4
               cAlingnment     =   6
               cButton         =   -1  'True
               cPositionY      =   20
               Begin Cis3D_v60.CIS3D CIS3D42 
                  Height          =   195
                  Left            =   495
                  Top             =   345
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   344
                  ForeColor       =   16711680
                  Caption         =   "(F12)"
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
            End
            Begin VB.ComboBox B1Comb_Bumon1 
               BackColor       =   &H00C0E0FF&
               Height          =   360
               Left            =   1020
               Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
               TabIndex        =   86
               Top             =   135
               Width           =   2745
            End
            Begin VB.ComboBox B1Comb_Bumon2 
               BackColor       =   &H00C0E0FF&
               Height          =   360
               Left            =   4755
               Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
               TabIndex        =   87
               Top             =   135
               Width           =   2745
            End
            Begin Cis3D_v60.CIS3D CIS3D39 
               Height          =   345
               Left            =   120
               Top             =   150
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   609
               ForeColor       =   16711680
               Caption         =   "ïîñÂá@"
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
               cAlingnment     =   7
               cBoderStyle     =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D40 
               Height          =   345
               Left            =   3855
               Top             =   150
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   609
               ForeColor       =   16711680
               Caption         =   "ïîñÂáA"
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
               cAlingnment     =   7
               cBoderStyle     =   2
            End
         End
      End
   End
   Begin Cis3D_v60.CIS3D Back_ChgTorcd 
      Height          =   2760
      Left            =   9990
      Top             =   5445
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4868
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
      cBoderWidth     =   3
      Begin Cis3D_v60.CIS3D CIS3D20 
         Height          =   2145
         Left            =   60
         Top             =   540
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   3784
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
         cBoderWidth     =   3
         cBoderStyle     =   1
         Begin vsFlexLib.vsFlexArray vsGrid1 
            Height          =   2070
            Left            =   45
            TabIndex        =   121
            TabStop         =   0   'False
            Top             =   30
            Width           =   4815
            _Version        =   196608
            _ExtentX        =   8493
            _ExtentY        =   3651
            _StockProps     =   228
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "NO|ïœä∑å≥CD|éÛì¸|Å@|é©é–éÛì¸|å_ñÒêÊ  |"
            Rows            =   6
            Cols            =   7
            BackColor       =   12648384
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   0
            RowHeightMin    =   330
            SelectionMode   =   1
            BackColorAlternate=   12648384
            Begin Cis3D_v60.CIS3D B1lb_Migiya 
               Height          =   345
               Left            =   1950
               Top             =   330
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   609
               BackColor       =   16777152
               Caption         =   "Å®"
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
            End
            Begin CisText_V60.CisText iB1_MotoUke 
               Height          =   345
               Left            =   1380
               TabIndex        =   61
               Top             =   330
               Width           =   585
               _ExtentX        =   1032
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
               Text            =   "XX"
               MaxLength       =   2
            End
            Begin CisText_V60.CisText iB1_ChgUke 
               Height          =   345
               Left            =   2250
               TabIndex        =   62
               Top             =   315
               Width           =   1095
               _ExtentX        =   1931
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
               Text            =   "XX"
               MaxLength       =   2
            End
            Begin CisText_V60.CisText iB1_MotoCD 
               Height          =   375
               Left            =   330
               TabIndex        =   122
               Top             =   300
               Width           =   1065
               _ExtentX        =   1879
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
            Begin CisText_V60.CisText iB1_HenKeiyaku 
               Height          =   375
               Left            =   3330
               TabIndex        =   63
               Top             =   300
               Width           =   1065
               _ExtentX        =   1879
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
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D19 
         Height          =   480
         Left            =   15
         Top             =   15
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   847
         BackColor       =   16761087
         ForeColor       =   192
         Caption         =   "[éÛíç]ïœä∑éÊà¯êÊÉRÅ[Éhìoò^"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
      End
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   9525
      Left            =   1425
      ScaleHeight     =   9465
      ScaleWidth      =   12480
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   1560
      Width           =   12540
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   9375
         Index           =   0
         Left            =   60
         Top             =   60
         Width           =   12390
         _ExtentX        =   21855
         _ExtentY        =   16536
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
         Begin Cis3D_v60.CIS3D Back_MK 
            Height          =   900
            Left            =   9630
            Top             =   3705
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   1588
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
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   840
               Left            =   30
               Top             =   30
               Width           =   2670
               _ExtentX        =   4710
               _ExtentY        =   1482
               ForeColor       =   16711680
               Caption         =   "[ëOçHíˆî≠çs]î≠íçéwé¶"
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
               Begin VB.CheckBox B1Chk_MaeHak 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "ï\é¶ñ≥"
                  Height          =   345
                  Left            =   780
                  TabIndex        =   65
                  Top             =   420
                  Width           =   1035
               End
            End
         End
         Begin Cis3D_v60.CIS3D Back_Info 
            Height          =   1890
            Left            =   60
            Top             =   5895
            Width           =   12285
            _ExtentX        =   21669
            _ExtentY        =   3334
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
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   1140
               Index           =   4
               Left            =   30
               Top             =   720
               Width           =   5445
               _ExtentX        =   9604
               _ExtentY        =   2011
               ForeColor       =   16711680
               Caption         =   "èZÅ@èä"
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
               Begin CisText_V60.CisText iB1_JAdd1 
                  Height          =   375
                  Left            =   315
                  TabIndex        =   70
                  Top             =   315
                  Width           =   4995
                  _ExtentX        =   8811
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
                  Text            =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
                  MaxLength       =   40
                  IMEMode         =   4
               End
               Begin CisText_V60.CisText iB1_JAdd2 
                  Height          =   375
                  Left            =   315
                  TabIndex        =   71
                  Top             =   675
                  Width           =   4995
                  _ExtentX        =   8811
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
                  Text            =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
                  MaxLength       =   40
                  IMEMode         =   4
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D12 
               Height          =   690
               Index           =   0
               Left            =   30
               Top             =   30
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   1217
               ForeColor       =   16711680
               Caption         =   "óXï÷î‘çÜ"
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
               cAlingnment     =   1
               cPositionX      =   100
               Begin CisText_V60.CisText iB1_JYubin1 
                  Height          =   375
                  Left            =   1095
                  TabIndex        =   66
                  Top             =   195
                  Width           =   480
                  _ExtentX        =   847
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
                  Text            =   "XXX"
                  MaxLength       =   3
                  cDataReplace    =   1
               End
               Begin CisText_V60.CisText iB1_JYubin2 
                  Height          =   375
                  Left            =   1560
                  TabIndex        =   67
                  Top             =   195
                  Width           =   600
                  _ExtentX        =   1058
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
                  Text            =   "XXXX"
                  MaxLength       =   4
                  cDataReplace    =   1
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D10 
               Height          =   450
               Index           =   1
               Left            =   5475
               Top             =   1410
               Width           =   6765
               _ExtentX        =   11933
               _ExtentY        =   794
               ForeColor       =   16711680
               Caption         =   "e-Mail"
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
               cAlingnment     =   1
               cPositionX      =   100
               Begin CisText_V60.CisText iB1_JTEMail 
                  Height          =   345
                  Left            =   1020
                  TabIndex        =   79
                  Top             =   60
                  Width           =   5385
                  _ExtentX        =   9499
                  _ExtentY        =   609
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
                  cDataReplace    =   2
                  cCaps           =   2
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D5 
               Height          =   480
               Index           =   1
               Left            =   5475
               Top             =   30
               Width           =   6765
               _ExtentX        =   11933
               _ExtentY        =   847
               ForeColor       =   16711680
               Caption         =   "ïîèêñº"
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
               cAlingnment     =   1
               cPositionX      =   100
               Begin CisText_V60.CisText iB1_JBusyo1 
                  Height          =   375
                  Left            =   1035
                  TabIndex        =   72
                  Top             =   30
                  Width           =   2580
                  _ExtentX        =   4551
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
                  Text            =   "XXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   20
                  IMEMode         =   4
               End
               Begin CisText_V60.CisText iB1_JBusyo2 
                  Height          =   375
                  Left            =   3600
                  TabIndex        =   73
                  Top             =   30
                  Width           =   2580
                  _ExtentX        =   4551
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
                  Text            =   "XXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   20
                  IMEMode         =   4
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D12 
               Height          =   450
               Index           =   1
               Left            =   5475
               Top             =   510
               Width           =   6765
               _ExtentX        =   11933
               _ExtentY        =   794
               ForeColor       =   16711680
               Caption         =   "ñêEñº"
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
               cAlingnment     =   1
               cPositionX      =   100
               Begin CisText_V60.CisText iB1_JYaku 
                  Height          =   375
                  Left            =   1035
                  TabIndex        =   74
                  Top             =   30
                  Width           =   2580
                  _ExtentX        =   4551
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
                  Text            =   "XXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   20
                  IMEMode         =   4
               End
               Begin CisText_V60.CisText iB1_JTanto 
                  Height          =   375
                  Left            =   4635
                  TabIndex        =   75
                  Top             =   30
                  Width           =   1815
                  _ExtentX        =   3201
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
                  Text            =   "XXXXXXXXXXXXXX"
                  MaxLength       =   14
                  IMEMode         =   4
               End
               Begin Cis3D_v60.CIS3D UC_3D15 
                  Height          =   375
                  Index           =   13
                  Left            =   3810
                  Top             =   45
                  Width           =   870
                  _ExtentX        =   1535
                  _ExtentY        =   661
                  ForeColor       =   16711680
                  Caption         =   "íSìñé“"
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   450
               Index           =   1
               Left            =   5475
               Top             =   960
               Width           =   4485
               _ExtentX        =   7911
               _ExtentY        =   794
               ForeColor       =   16711680
               Caption         =   "Tel."
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
               cAlingnment     =   1
               cPositionX      =   100
               Begin CisText_V60.CisText iB1_JTTel1 
                  Height          =   345
                  Left            =   1020
                  TabIndex        =   76
                  Top             =   60
                  Width           =   1725
                  _ExtentX        =   3043
                  _ExtentY        =   609
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XXXXXXXXXXXXX"
                  MaxLength       =   13
               End
               Begin CisText_V60.CisText iB1_JTTel2 
                  Height          =   345
                  Left            =   2700
                  TabIndex        =   77
                  Top             =   60
                  Width           =   1725
                  _ExtentX        =   3043
                  _ExtentY        =   609
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XXXXXXXXXXXXX"
                  MaxLength       =   13
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D9 
               Height          =   450
               Index           =   1
               Left            =   9960
               Top             =   960
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   794
               ForeColor       =   16711680
               Caption         =   "Fax."
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
               cAlingnment     =   1
               cPositionX      =   100
               Begin CisText_V60.CisText iB1_JTFax 
                  Height          =   375
                  Left            =   585
                  TabIndex        =   78
                  Top             =   60
                  Width           =   1590
                  _ExtentX        =   2805
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
                  Text            =   "XXXXXXXXX1XX"
                  MaxLength       =   12
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   690
               Index           =   0
               Left            =   2325
               Top             =   30
               Width           =   3150
               _ExtentX        =   5556
               _ExtentY        =   1217
               ForeColor       =   16711680
               Caption         =   "ë„ï\"
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
               cAlingnment     =   1
               cPositionX      =   200
               Begin CisText_V60.CisText iB1_JTel 
                  Height          =   345
                  Left            =   1350
                  TabIndex        =   68
                  Top             =   15
                  Width           =   1605
                  _ExtentX        =   2831
                  _ExtentY        =   609
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XXXXXXXXX1XX"
                  MaxLength       =   12
               End
               Begin CisText_V60.CisText iB1_JFax 
                  Height          =   345
                  Left            =   1335
                  TabIndex        =   69
                  Top             =   345
                  Width           =   1620
                  _ExtentX        =   2858
                  _ExtentY        =   609
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XXXXXXXXX1XX"
                  MaxLength       =   12
               End
               Begin Cis3D_v60.CIS3D UC_3D9 
                  Height          =   285
                  Index           =   0
                  Left            =   795
                  Top             =   375
                  Width           =   450
                  _ExtentX        =   794
                  _ExtentY        =   503
                  ForeColor       =   16711680
                  Caption         =   "Fax."
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
                  cAlingnment     =   1
                  cBoderStyle     =   2
               End
               Begin Cis3D_v60.CIS3D UC_3D9 
                  Height          =   285
                  Index           =   2
                  Left            =   780
                  Top             =   30
                  Width           =   450
                  _ExtentX        =   794
                  _ExtentY        =   503
                  ForeColor       =   16711680
                  Caption         =   "Tel."
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
                  cAlingnment     =   1
                  cBoderStyle     =   2
               End
            End
         End
         Begin Cis3D_v60.CIS3D Back_HC 
            Height          =   900
            Left            =   9630
            Top             =   3705
            Visible         =   0   'False
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   1588
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
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   840
               Left            =   30
               Top             =   30
               Width           =   2670
               _ExtentX        =   4710
               _ExtentY        =   1482
               ForeColor       =   16711680
               Caption         =   "äJénî≠íçì˙"
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
               Begin CisYMD_With_Btn.CisYMDwB iB1_STHYmd 
                  Height          =   345
                  Left            =   360
                  TabIndex        =   64
                  Top             =   390
                  Width           =   1890
                  _ExtentX        =   3069
                  _ExtentY        =   556
                  cTextColor      =   0
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
                  Object.Width           =   1890
                  cSize           =   1
                  cChkResult      =   0   'False
                  cTypeIntval     =   0
                  BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
            End
         End
         Begin Cis3D_v60.CIS3D Back_Send 
            Height          =   1545
            Left            =   60
            Top             =   7785
            Width           =   12285
            _ExtentX        =   21669
            _ExtentY        =   2725
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
            Begin Cis3D_v60.CIS3D UC_3D16 
               Height          =   1170
               Index           =   3
               Left            =   4695
               Top             =   330
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   2064
               BackColor       =   16761024
               ForeColor       =   0
               Caption         =   "e-Mail"
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
               cFont3D         =   2
               cAlingnment     =   7
               cPositionX      =   30
            End
            Begin Cis3D_v60.CIS3D CIS3D3 
               Height          =   1170
               Index           =   0
               Left            =   345
               Top             =   330
               Width           =   4350
               _ExtentX        =   7673
               _ExtentY        =   2064
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
               Begin Cis3D_v60.CIS3D B1lb_SendFax 
                  Height          =   345
                  Index           =   1
                  Left            =   2760
                  Top             =   405
                  Width           =   1530
                  _ExtentX        =   2699
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
               Begin Cis3D_v60.CIS3D B1lb_SendFax 
                  Height          =   345
                  Index           =   0
                  Left            =   2760
                  Top             =   45
                  Width           =   1530
                  _ExtentX        =   2699
                  _ExtentY        =   609
                  BackColor       =   16777152
                  Caption         =   "0566-12-1234"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
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
               Begin CisText_V60.CisText iB1_SendFax 
                  Height          =   345
                  Index           =   0
                  Left            =   30
                  TabIndex        =   80
                  Top             =   45
                  Width           =   975
                  _ExtentX        =   1720
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
               Begin Cis3D_v60.CIS3D B1lb_SendFaxTor 
                  Height          =   345
                  Index           =   0
                  Left            =   990
                  Top             =   45
                  Width           =   1755
                  _ExtentX        =   3096
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
               Begin CisText_V60.CisText iB1_SendFax 
                  Height          =   360
                  Index           =   1
                  Left            =   30
                  TabIndex        =   81
                  Top             =   390
                  Width           =   975
                  _ExtentX        =   1720
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
               Begin Cis3D_v60.CIS3D B1lb_SendFaxTor 
                  Height          =   345
                  Index           =   1
                  Left            =   990
                  Top             =   405
                  Width           =   1755
                  _ExtentX        =   3096
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
               Begin CisText_V60.CisText iB1_SendFax 
                  Height          =   360
                  Index           =   2
                  Left            =   30
                  TabIndex        =   82
                  Top             =   750
                  Width           =   975
                  _ExtentX        =   1720
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
               Begin Cis3D_v60.CIS3D B1lb_SendFaxTor 
                  Height          =   345
                  Index           =   2
                  Left            =   990
                  Top             =   765
                  Width           =   1755
                  _ExtentX        =   3096
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
               Begin Cis3D_v60.CIS3D B1lb_SendFax 
                  Height          =   345
                  Index           =   2
                  Left            =   2760
                  Top             =   765
                  Width           =   1530
                  _ExtentX        =   2699
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
            Begin Cis3D_v60.CIS3D UC_3D16 
               Height          =   315
               Index           =   0
               Left            =   30
               Top             =   15
               Width           =   12225
               _ExtentX        =   21564
               _ExtentY        =   556
               BackColor       =   16761024
               ForeColor       =   0
               Caption         =   "Éf Å[ É^ ëó êM êÊ í« â¡ èÓ ïÒ"
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
               cFont3D         =   2
               cAlingnment     =   7
               cPositionX      =   50
            End
            Begin Cis3D_v60.CIS3D UC_3D16 
               Height          =   1170
               Index           =   1
               Left            =   30
               Top             =   330
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   2064
               BackColor       =   16761024
               ForeColor       =   0
               Caption         =   "ÇeÇ`Çw"
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
               cFont3D         =   2
               cPositionX      =   30
               cPositionY      =   200
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   1170
               Left            =   5445
               Top             =   330
               Width           =   6810
               _ExtentX        =   12012
               _ExtentY        =   2064
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
               Begin Cis3D_v60.CIS3D B1lb_SendMail 
                  Height          =   345
                  Index           =   1
                  Left            =   2790
                  Top             =   405
                  Width           =   3960
                  _ExtentX        =   6985
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
               Begin Cis3D_v60.CIS3D B1lb_SendMail 
                  Height          =   345
                  Index           =   0
                  Left            =   2790
                  Top             =   45
                  Width           =   3960
                  _ExtentX        =   6985
                  _ExtentY        =   609
                  BackColor       =   16777152
                  Caption         =   "nam-cis@cisweb.co.jp"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
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
               Begin CisText_V60.CisText iB1_SendMail 
                  Height          =   345
                  Index           =   0
                  Left            =   60
                  TabIndex        =   83
                  Top             =   45
                  Width           =   975
                  _ExtentX        =   1720
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
               Begin Cis3D_v60.CIS3D B1lb_SendMailTor 
                  Height          =   345
                  Index           =   0
                  Left            =   1020
                  Top             =   45
                  Width           =   1755
                  _ExtentX        =   3096
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
               Begin CisText_V60.CisText iB1_SendMail 
                  Height          =   360
                  Index           =   1
                  Left            =   60
                  TabIndex        =   84
                  Top             =   390
                  Width           =   975
                  _ExtentX        =   1720
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
               Begin Cis3D_v60.CIS3D B1lb_SendMailTor 
                  Height          =   345
                  Index           =   1
                  Left            =   1020
                  Top             =   405
                  Width           =   1755
                  _ExtentX        =   3096
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
               Begin CisText_V60.CisText iB1_SendMail 
                  Height          =   360
                  Index           =   2
                  Left            =   60
                  TabIndex        =   85
                  Top             =   750
                  Width           =   975
                  _ExtentX        =   1720
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
               Begin Cis3D_v60.CIS3D B1lb_SendMailTor 
                  Height          =   345
                  Index           =   2
                  Left            =   1020
                  Top             =   765
                  Width           =   1755
                  _ExtentX        =   3096
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
               Begin Cis3D_v60.CIS3D B1lb_SendMail 
                  Height          =   345
                  Index           =   2
                  Left            =   2790
                  Top             =   765
                  Width           =   3960
                  _ExtentX        =   6985
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
         End
         Begin Cis3D_v60.CIS3D UC_3D4 
            Height          =   3660
            Left            =   60
            Top             =   30
            Width           =   12300
            _ExtentX        =   21696
            _ExtentY        =   6456
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
            Begin Cis3D_v60.CIS3D Back_Syukaba 
               Height          =   405
               Left            =   4605
               Top             =   3225
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   714
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
               Begin CisText_V60.CisText iB1_Syukaba 
                  Height          =   375
                  Left            =   1080
                  TabIndex        =   46
                  Top             =   15
                  Width           =   480
                  _ExtentX        =   847
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
                  Text            =   "XXX"
                  MaxLength       =   3
               End
               Begin Cis3D_v60.CIS3D CIS3D22 
                  Height          =   285
                  Left            =   90
                  Top             =   30
                  Width           =   735
                  _ExtentX        =   1296
                  _ExtentY        =   503
                  ForeColor       =   16711680
                  Caption         =   "èoâ◊èÍ"
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
            End
            Begin Cis3D_v60.CIS3D Back_SJigyo 
               Height          =   405
               Left            =   2400
               Top             =   3225
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   714
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
               Begin CisText_V60.CisText iB1_SJigyoCD 
                  Height          =   375
                  Left            =   1845
                  TabIndex        =   45
                  Top             =   15
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
               Begin Cis3D_v60.CIS3D CIS3D21 
                  Height          =   285
                  Left            =   135
                  Top             =   60
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   503
                  ForeColor       =   16711680
                  Caption         =   "édì¸êÊéñã∆ëÃ"
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
            End
            Begin Cis3D_v60.CIS3D Back_NJigyo 
               Height          =   405
               Left            =   45
               Top             =   3225
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   714
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
               Begin CisText_V60.CisText iB1_NJigyoCD 
                  Height          =   375
                  Left            =   1755
                  TabIndex        =   43
                  Top             =   15
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
               Begin Cis3D_v60.CIS3D CIS3D18 
                  Height          =   285
                  Left            =   120
                  Top             =   60
                  Width           =   1500
                  _ExtentX        =   2646
                  _ExtentY        =   503
                  ForeColor       =   16711680
                  Caption         =   "î[ì¸êÊéñã∆ëÃ"
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
               Begin CisText_V60.CisText iB1_NJigyoKJ 
                  Height          =   375
                  Left            =   2085
                  TabIndex        =   44
                  Top             =   15
                  Width           =   240
                  _ExtentX        =   423
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
                  Text            =   "X"
                  MaxLength       =   1
               End
            End
            Begin Cis3D_v60.CIS3D Back_HTorcd 
               Height          =   690
               Left            =   9135
               Top             =   30
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   1217
               ForeColor       =   16711680
               Caption         =   "îÌé©é–CD"
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
               Begin CisText_V60.CisText iB1_HTorcd 
                  Height          =   375
                  Left            =   240
                  TabIndex        =   9
                  Top             =   285
                  Width           =   735
                  _ExtentX        =   1296
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
                  Text            =   "XXXXX"
                  MaxLength       =   5
               End
               Begin CisText_V60.CisText iB1_HKojyo 
                  Height          =   375
                  Left            =   975
                  TabIndex        =   10
                  Top             =   285
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
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   690
               Index           =   0
               Left            =   2355
               Top             =   30
               Width           =   6780
               _ExtentX        =   11959
               _ExtentY        =   1217
               ForeColor       =   16711680
               Caption         =   "ê≥éÆñºèÃ"
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
               Begin CisText_V60.CisText iB1_Tornm 
                  Height          =   375
                  Left            =   285
                  TabIndex        =   6
                  Top             =   300
                  Width           =   6240
                  _ExtentX        =   11007
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
                  Text            =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4XXXXXXXXX5"
                  MaxLength       =   50
                  IMEMode         =   4
               End
            End
            Begin Cis3D_v60.CIS3D Back_ConvTorcd1 
               Height          =   690
               Left            =   8280
               Top             =   30
               Visible         =   0   'False
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   1217
               ForeColor       =   16711680
               Caption         =   "ïœä∑å≥éÊà¯êÊCD"
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
               Begin CisText_V60.CisText iB1_ConvTorcd 
                  Height          =   375
                  Left            =   330
                  TabIndex        =   7
                  Top             =   300
                  Width           =   735
                  _ExtentX        =   1296
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
                  Text            =   "XXXXX"
                  MaxLength       =   5
               End
               Begin CisText_V60.CisText iB1_ConvTorcdKj 
                  Height          =   375
                  Left            =   1065
                  TabIndex        =   8
                  Top             =   300
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
            Begin Cis3D_v60.CIS3D Back_ConvTorcd2 
               Height          =   690
               Left            =   8280
               Top             =   30
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   1217
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
            Begin Cis3D_v60.CIS3D CIS3D7 
               Height          =   720
               Left            =   4725
               Top             =   2505
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   1270
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
               Begin VB.CheckBox B1Chk_PrtkbJ 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "î≠çsóL"
                  Height          =   300
                  Left            =   195
                  TabIndex        =   38
                  Top             =   360
                  Width           =   1035
               End
            End
            Begin Cis3D_v60.CIS3D Back_Juryo 
               Height          =   720
               Left            =   6180
               Top             =   2505
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   1270
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
               Begin VB.CheckBox B1Chk_Juryo 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "èoâ◊éû"
                  Height          =   240
                  Left            =   225
                  TabIndex        =   39
                  Top             =   375
                  Value           =   1  '¡™Ø∏
                  Width           =   1095
               End
            End
            Begin Cis3D_v60.CIS3D Back_KojyoNm 
               Height          =   690
               Left            =   5145
               Top             =   720
               Width           =   2700
               _ExtentX        =   4763
               _ExtentY        =   1217
               ForeColor       =   16711680
               Caption         =   "çHèÍñºèÃ"
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
               Begin CisText_V60.CisText iB1_KojyoNm 
                  Height          =   375
                  Left            =   90
                  TabIndex        =   14
                  Top             =   285
                  Width           =   2565
                  _ExtentX        =   4524
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
                  Text            =   "XXXXXXXXX1XXXXXXXXX2"
                  MaxLength       =   20
                  IMEMode         =   4
               End
            End
            Begin Cis3D_v60.CIS3D Back_KTorcd 
               Height          =   690
               Left            =   7845
               Top             =   720
               Width           =   2955
               _ExtentX        =   5212
               _ExtentY        =   1217
               ForeColor       =   16711680
               Caption         =   "å_Å@ñÒ  êÊ"
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
               Begin CisText_V60.CisText iB1_KTorcd 
                  Height          =   360
                  Left            =   60
                  TabIndex        =   15
                  Top             =   285
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
               Begin Cis3D_v60.CIS3D B1lb_KTorNm 
                  Height          =   360
                  Left            =   1050
                  Top             =   285
                  Width           =   1785
                  _ExtentX        =   3149
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
            Begin Cis3D_v60.CIS3D Back_Keisan 
               Height          =   1095
               Left            =   45
               Top             =   1410
               Width           =   8220
               _ExtentX        =   14499
               _ExtentY        =   1931
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
               cPositionX      =   30
               Begin Cis3D_v60.CIS3D CIS3D12 
                  Height          =   390
                  Left            =   15
                  Top             =   180
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   688
                  ForeColor       =   16711680
                  Caption         =   "åvéZ"
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
                  cAlingnment     =   7
                  cBoderStyle     =   2
               End
               Begin Cis3D_v60.CIS3D CIS3D13 
                  Height          =   390
                  Left            =   15
                  Top             =   510
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   688
                  ForeColor       =   16711680
                  Caption         =   "ï˚ñ@"
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
                  cAlingnment     =   7
                  cBoderStyle     =   2
               End
               Begin Cis3D_v60.CIS3D Back_KeisanH 
                  Height          =   960
                  Left            =   2055
                  Top             =   75
                  Width           =   6135
                  _ExtentX        =   10821
                  _ExtentY        =   1693
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
                  Begin VB.OptionButton B1Op_Marume 
                     BackColor       =   &H00C0C0C0&
                     Caption         =   "éléÃå‹ì¸"
                     Height          =   285
                     Index           =   2
                     Left            =   120
                     TabIndex        =   21
                     Top             =   645
                     Width           =   1275
                  End
                  Begin VB.OptionButton B1Op_Marume 
                     BackColor       =   &H00C0C0C0&
                     Caption         =   "êÿÇËéÃÇƒ"
                     Height          =   285
                     Index           =   0
                     Left            =   120
                     TabIndex        =   19
                     Top             =   30
                     Width           =   1335
                  End
                  Begin VB.OptionButton B1Op_Marume 
                     BackColor       =   &H00C0C0C0&
                     Caption         =   "êÿè„Ç∞"
                     Height          =   300
                     Index           =   1
                     Left            =   120
                     TabIndex        =   20
                     Top             =   330
                     Width           =   1065
                  End
                  Begin VB.OptionButton B1Op_Marume 
                     BackColor       =   &H00C0C0C0&
                     Caption         =   "è⁄ç◊ê›íË"
                     Height          =   600
                     Index           =   3
                     Left            =   1455
                     TabIndex        =   22
                     Top             =   75
                     Width           =   825
                  End
                  Begin Cis3D_v60.CIS3D CIS3D17 
                     Height          =   225
                     Left            =   1500
                     Top             =   720
                     Width           =   4665
                     _ExtentX        =   8229
                     _ExtentY        =   397
                     ForeColor       =   255
                     Caption         =   "Å¶ è⁄ç◊ê›íËÇÕé¿ê—èàóùÇ…ÇÃÇ›îΩâfÇ≥ÇÍÇ‹Ç∑"
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
                     cFont3D         =   2
                     cBoderStyle     =   2
                  End
                  Begin Cis3D_v60.CIS3D Back_Kurai 
                     Height          =   690
                     Left            =   2280
                     Top             =   30
                     Width           =   3825
                     _ExtentX        =   6747
                     _ExtentY        =   1217
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
                     Begin Cis3D_v60.CIS3D CIS3D10 
                        Height          =   315
                        Left            =   1005
                        Top             =   15
                        Width           =   2805
                        _ExtentX        =   4948
                        _ExtentY        =   556
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
                        cBoderColor1    =   12632256
                        cBoderColor2    =   12632256
                        cFont3DColor1   =   16777215
                        cFont3DColor2   =   8421504
                        cBoderStyle     =   2
                        Begin VB.OptionButton B1Op_Kurai1 
                           BackColor       =   &H00C0C0C0&
                           Caption         =   "éléÃå‹ì¸"
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
                           Left            =   1620
                           TabIndex        =   25
                           Top             =   45
                           Width           =   1140
                        End
                        Begin VB.OptionButton B1Op_Kurai1 
                           BackColor       =   &H00C0C0C0&
                           Caption         =   "êÿè„"
                           BeginProperty Font 
                              Name            =   "ÇlÇr ÉSÉVÉbÉN"
                              Size            =   9.75
                              Charset         =   128
                              Weight          =   400
                              Underline       =   0   'False
                              Italic          =   0   'False
                              Strikethrough   =   0   'False
                           EndProperty
                           Height          =   240
                           Index           =   1
                           Left            =   900
                           TabIndex        =   24
                           Top             =   30
                           Width           =   735
                        End
                        Begin VB.OptionButton B1Op_Kurai1 
                           BackColor       =   &H00C0C0C0&
                           Caption         =   "êÿéÃ"
                           BeginProperty Font 
                              Name            =   "ÇlÇr ÉSÉVÉbÉN"
                              Size            =   9.75
                              Charset         =   128
                              Weight          =   400
                              Underline       =   0   'False
                              Italic          =   0   'False
                              Strikethrough   =   0   'False
                           EndProperty
                           Height          =   240
                           Index           =   0
                           Left            =   135
                           TabIndex        =   23
                           Top             =   30
                           Width           =   735
                        End
                     End
                     Begin Cis3D_v60.CIS3D CIS3D8 
                        Height          =   315
                        Left            =   990
                        Top             =   360
                        Width           =   2820
                        _ExtentX        =   4974
                        _ExtentY        =   556
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
                        cBoderColor1    =   12632256
                        cBoderColor2    =   12632256
                        cFont3DColor1   =   16777215
                        cFont3DColor2   =   8421504
                        cBoderStyle     =   2
                        Begin VB.OptionButton B1Op_Kurai2 
                           BackColor       =   &H00C0C0C0&
                           Caption         =   "êÿéÃ"
                           BeginProperty Font 
                              Name            =   "ÇlÇr ÉSÉVÉbÉN"
                              Size            =   9.75
                              Charset         =   128
                              Weight          =   400
                              Underline       =   0   'False
                              Italic          =   0   'False
                              Strikethrough   =   0   'False
                           EndProperty
                           Height          =   240
                           Index           =   0
                           Left            =   150
                           TabIndex        =   26
                           Top             =   45
                           Width           =   735
                        End
                        Begin VB.OptionButton B1Op_Kurai2 
                           BackColor       =   &H00C0C0C0&
                           Caption         =   "êÿè„"
                           BeginProperty Font 
                              Name            =   "ÇlÇr ÉSÉVÉbÉN"
                              Size            =   9.75
                              Charset         =   128
                              Weight          =   400
                              Underline       =   0   'False
                              Italic          =   0   'False
                              Strikethrough   =   0   'False
                           EndProperty
                           Height          =   240
                           Index           =   1
                           Left            =   915
                           TabIndex        =   27
                           Top             =   45
                           Width           =   735
                        End
                        Begin VB.OptionButton B1Op_Kurai2 
                           BackColor       =   &H00C0C0C0&
                           Caption         =   "éléÃå‹ì¸"
                           BeginProperty Font 
                              Name            =   "ÇlÇr ÉSÉVÉbÉN"
                              Size            =   9.75
                              Charset         =   128
                              Weight          =   400
                              Underline       =   0   'False
                              Italic          =   0   'False
                              Strikethrough   =   0   'False
                           EndProperty
                           Height          =   240
                           Index           =   2
                           Left            =   1635
                           TabIndex        =   28
                           Top             =   45
                           Width           =   1140
                        End
                     End
                     Begin Cis3D_v60.CIS3D CIS3D14 
                        Height          =   225
                        Left            =   45
                        Top             =   30
                        Width           =   960
                        _ExtentX        =   1693
                        _ExtentY        =   397
                        ForeColor       =   0
                        Caption         =   "è¨êîì_1åÖ"
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
                        cFont3D         =   2
                        cBoderStyle     =   2
                     End
                     Begin Cis3D_v60.CIS3D CIS3D16 
                        Height          =   225
                        Left            =   45
                        Top             =   390
                        Width           =   960
                        _ExtentX        =   1693
                        _ExtentY        =   397
                        ForeColor       =   0
                        Caption         =   "è¨êîì_2åÖ"
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
                        cFont3D         =   2
                        cBoderStyle     =   2
                     End
                     Begin VB.Line Line1 
                        Index           =   0
                        X1              =   30
                        X2              =   3780
                        Y1              =   330
                        Y2              =   330
                     End
                     Begin VB.Line Line1 
                        BorderColor     =   &H00FFFFFF&
                        Index           =   1
                        X1              =   30
                        X2              =   3780
                        Y1              =   345
                        Y2              =   345
                     End
                  End
               End
               Begin Cis3D_v60.CIS3D Back_KeisanT 
                  Height          =   960
                  Left            =   660
                  Top             =   75
                  Width           =   1380
                  _ExtentX        =   2434
                  _ExtentY        =   1693
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
                  Begin VB.OptionButton B1Op_Keisan 
                     BackColor       =   &H00C0C0C0&
                     Caption         =   "åéíPà "
                     Height          =   345
                     Index           =   1
                     Left            =   60
                     TabIndex        =   18
                     Top             =   510
                     Width           =   1035
                  End
                  Begin VB.OptionButton B1Op_Keisan 
                     BackColor       =   &H00FFC0C0&
                     Caption         =   "ì`ï[íPà "
                     Height          =   345
                     Index           =   0
                     Left            =   60
                     TabIndex        =   17
                     Top             =   165
                     Width           =   1305
                  End
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   690
               Index           =   1
               Left            =   45
               Top             =   720
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   1217
               ForeColor       =   16711680
               Caption         =   "çıà¯∂≈"
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
               Begin CisText_V60.CisText iB1_SKana 
                  Height          =   330
                  Left            =   330
                  TabIndex        =   12
                  Top             =   285
                  Width           =   240
                  _ExtentX        =   423
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
                  Text            =   "X"
                  MaxLength       =   1
                  IMEMode         =   6
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   690
               Index           =   7
               Left            =   45
               Top             =   30
               Width           =   2310
               _ExtentX        =   4075
               _ExtentY        =   1217
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
               Begin VB.CheckBox B1Chk_Torkb 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "é–ì‡"
                  Height          =   240
                  Index           =   1
                  Left            =   1260
                  TabIndex        =   5
                  Top             =   225
                  Width           =   810
               End
               Begin VB.CheckBox B1Chk_Torkb 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "å_ñÒêÊ"
                  Height          =   240
                  Index           =   0
                  Left            =   180
                  TabIndex        =   4
                  Top             =   225
                  Width           =   1095
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D13 
               Height          =   690
               Index           =   0
               Left            =   945
               Top             =   720
               Width           =   4200
               _ExtentX        =   7408
               _ExtentY        =   1217
               ForeColor       =   16711680
               Caption         =   " ó™èÃ"
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
               cPositionX      =   -700
               Begin CisText_V60.CisText iB1_Ryaku 
                  Height          =   375
                  Left            =   90
                  TabIndex        =   13
                  Top             =   285
                  Width           =   2565
                  _ExtentX        =   4524
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
                  Text            =   "XXXXXXXXX1XXXXXXXXX2"
                  MaxLength       =   20
                  IMEMode         =   4
               End
               Begin Cis3D_v60.CIS3D B1lb_Ryaku1 
                  Height          =   330
                  Left            =   2730
                  Top             =   15
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   582
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
               Begin Cis3D_v60.CIS3D B1lb_Ryaku2 
                  Height          =   330
                  Left            =   2730
                  Top             =   345
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   582
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
            Begin Cis3D_v60.CIS3D Back_Taisyo 
               Height          =   690
               Left            =   10800
               Top             =   720
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   1217
               ForeColor       =   16711680
               Caption         =   " â∫êøñ@"
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
               Begin VB.CheckBox B1Chk_Taisyo 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "ëŒè€äO"
                  Height          =   255
                  Left            =   210
                  TabIndex        =   16
                  Top             =   330
                  Width           =   1050
               End
            End
            Begin Cis3D_v60.CIS3D Back_DenSyu 
               Height          =   720
               Left            =   45
               Top             =   2505
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   1270
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
               Begin CisText_V60.CisText iB1_DenSyu 
                  Height          =   345
                  Left            =   120
                  TabIndex        =   35
                  Top             =   315
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
               Begin Cis3D_v60.CIS3D B1lb_DenSyu 
                  Height          =   330
                  Left            =   360
                  Top             =   315
                  Width           =   1290
                  _ExtentX        =   2275
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
            Begin Cis3D_v60.CIS3D Back_Shime 
               Height          =   1095
               Left            =   8265
               Top             =   1410
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   1931
               ForeColor       =   16711680
               Caption         =   "í˜Çﬂì˙"
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
               cPositionX      =   10
               cPositionY      =   30
               Begin CisText_V60.CisText iB1_Shime 
                  Height          =   375
                  Left            =   165
                  TabIndex        =   29
                  Top             =   480
                  Width           =   420
                  _ExtentX        =   741
                  _ExtentY        =   661
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
            Begin Cis3D_v60.CIS3D Back_Keisu 
               Height          =   1095
               Left            =   11625
               Top             =   1410
               Width           =   630
               _ExtentX        =   1111
               _ExtentY        =   1931
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
               Begin CisText_V60.CisText iB1_Keisu 
                  Height          =   330
                  Left            =   30
                  TabIndex        =   34
                  Top             =   495
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
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
               Begin VB.Label Label2 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "%"
                  Height          =   315
                  Left            =   450
                  TabIndex        =   119
                  Top             =   555
                  Width           =   165
               End
            End
            Begin Cis3D_v60.CIS3D Back_Hei 
               Height          =   1095
               Left            =   9015
               Top             =   1410
               Width           =   2610
               _ExtentX        =   4604
               _ExtentY        =   1931
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
               Begin Cis3D_v60.CIS3D Back_Heijyun 
                  Height          =   405
                  Left            =   45
                  Top             =   465
                  Width           =   2535
                  _ExtentX        =   4471
                  _ExtentY        =   714
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
                  Begin VB.OptionButton B1Op_Hei 
                     BackColor       =   &H00C0C0C0&
                     Caption         =   "ñ≥"
                     Height          =   345
                     Index           =   0
                     Left            =   60
                     TabIndex        =   30
                     Top             =   45
                     Width           =   615
                  End
                  Begin VB.OptionButton B1Op_Hei 
                     BackColor       =   &H00C0C0C0&
                     Caption         =   "ì˙"
                     Height          =   345
                     Index           =   1
                     Left            =   675
                     TabIndex        =   31
                     Top             =   45
                     Width           =   615
                  End
                  Begin VB.OptionButton B1Op_Hei 
                     BackColor       =   &H00C0C0C0&
                     Caption         =   "åé"
                     Height          =   345
                     Index           =   3
                     Left            =   1890
                     TabIndex        =   33
                     Top             =   30
                     Width           =   615
                  End
                  Begin VB.OptionButton B1Op_Hei 
                     BackColor       =   &H00C0C0C0&
                     Caption         =   "ï÷"
                     Height          =   345
                     Index           =   2
                     Left            =   1260
                     TabIndex        =   32
                     Top             =   30
                     Width           =   615
                  End
               End
            End
            Begin Cis3D_v60.CIS3D Back_HNaiji 
               Height          =   720
               Left            =   9570
               Top             =   2505
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   1270
               ForeColor       =   16711680
               Caption         =   "î≠íçì‡é¶"
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
               Begin VB.ComboBox B1Comb_HNaiji 
                  BackColor       =   &H00C0E0FF&
                  Height          =   360
                  Left            =   105
                  Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
                  TabIndex        =   41
                  Top             =   315
                  Width           =   1365
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   720
               Index           =   5
               Left            =   11130
               Top             =   2505
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   1270
               ForeColor       =   16711680
               Caption         =   "ëOì|ì˙êî"
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
               Begin CisText_V60.CisText iB1_Mae 
                  Height          =   330
                  Left            =   330
                  TabIndex        =   42
                  Top             =   315
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
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D Back_Meaisai 
               Height          =   720
               Left            =   7695
               Top             =   2505
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   1270
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
               Begin VB.ComboBox B1Comb_Meaisai 
                  BackColor       =   &H00C0E0FF&
                  Height          =   360
                  Left            =   105
                  Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
                  TabIndex        =   40
                  Top             =   315
                  Width           =   1695
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   690
               Index           =   2
               Left            =   10710
               Top             =   30
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   1217
               ForeColor       =   16711680
               Caption         =   "çÜï‚ãÊï™"
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
               Begin VB.CheckBox B1Chk_Gouho 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "çÜå˚"
                  Height          =   345
                  Left            =   360
                  TabIndex        =   11
                  Top             =   315
                  Width           =   795
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D2 
               Height          =   720
               Left            =   1815
               Top             =   2505
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   1270
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
               Begin VB.CheckBox B1Chk_PrtkbN 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "î≠çsóL"
                  Height          =   300
                  Left            =   165
                  TabIndex        =   36
                  Top             =   360
                  Width           =   1035
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   720
               Left            =   3270
               Top             =   2505
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   1270
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
               Begin VB.CheckBox B1Chk_PrtkbS 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "î≠çsóL"
                  Height          =   300
                  Left            =   180
                  TabIndex        =   37
                  Top             =   360
                  Width           =   1035
               End
            End
            Begin Cis3D_v60.CIS3D Back_Maekotei 
               Height          =   405
               Index           =   0
               Left            =   45
               Top             =   3225
               Width           =   6135
               _ExtentX        =   10821
               _ExtentY        =   714
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
               Begin Cis3D_v60.CIS3D Back_Maekotei 
                  Height          =   405
                  Index           =   1
                  Left            =   2355
                  Top             =   0
                  Width           =   2205
                  _ExtentX        =   3889
                  _ExtentY        =   714
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
            End
            Begin Cis3D_v60.CIS3D CIS3D23 
               Height          =   405
               Left            =   6180
               Top             =   3240
               Width           =   6075
               _ExtentX        =   10716
               _ExtentY        =   714
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
               Begin Cis3D_v60.CIS3D Back_Sijisyo 
                  Height          =   390
                  Left            =   0
                  Top             =   0
                  Width           =   3390
                  _ExtentX        =   5980
                  _ExtentY        =   688
                  ForeColor       =   16711680
                  Caption         =   "èoâ◊éwé¶èëî≠çs"
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
                  Begin VB.CheckBox B1Chk_Sijisyo 
                     BackColor       =   &H00C0C0C0&
                     Caption         =   "î≠çsóL"
                     Height          =   300
                     Left            =   2085
                     TabIndex        =   47
                     Top             =   60
                     Width           =   1035
                  End
               End
               Begin Cis3D_v60.CIS3D Back_HinKeta 
                  Height          =   405
                  Left            =   3390
                  Top             =   0
                  Width           =   2685
                  _ExtentX        =   4736
                  _ExtentY        =   714
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
                  Begin CisText_V60.CisText iB1_HinKeta 
                     Height          =   375
                     Left            =   2205
                     TabIndex        =   48
                     Top             =   15
                     Width           =   330
                     _ExtentX        =   582
                     _ExtentY        =   661
                     cFormat         =   "##"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   11.25
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
                     cGFormat        =   "##"
                     cILength        =   2
                  End
                  Begin Cis3D_v60.CIS3D CIS3D43 
                     Height          =   285
                     Left            =   120
                     Top             =   60
                     Width           =   2055
                     _ExtentX        =   3625
                     _ExtentY        =   503
                     ForeColor       =   16711680
                     Caption         =   "ïiî‘É`ÉFÉbÉNåÖêî"
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
               End
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   2160
            Left            =   60
            Top             =   3720
            Width           =   9525
            _ExtentX        =   16801
            _ExtentY        =   3810
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
            Begin vsFlexLib.vsFlexArray vsGrid 
               Height          =   2085
               Left            =   195
               TabIndex        =   120
               TabStop         =   0   'False
               Top             =   30
               Width           =   9390
               _Version        =   196608
               _ExtentX        =   16563
               _ExtentY        =   3678
               _StockProps     =   228
               Appearance      =   1
               ConvInfo        =   1418783674
               FormatString    =   "NO|éÛ|éÛì¸ñºèÃ             |ÉTÉCÉNÉã|ãÛñæç◊|ójì˙î≠íç|XXXX|XXXX|XXXX|XXXX|XXXX|CYCLE1|CYCLE2|CYCLE3|ójì˙||"
               Rows            =   6
               Cols            =   21
               BackColor       =   12648384
               BackColorBkg    =   12632256
               FocusRect       =   0
               HighLight       =   0
               RowHeightMin    =   330
               SelectionMode   =   1
               BackColorAlternate=   12648384
               Begin VB.CheckBox ChkB1_Youbi 
                  BackColor       =   &H00C0E0FF&
                  Height          =   315
                  Index           =   4
                  Left            =   8460
                  TabIndex        =   60
                  Top             =   330
                  Width           =   555
               End
               Begin VB.CheckBox ChkB1_Youbi 
                  BackColor       =   &H00C0E0FF&
                  Height          =   315
                  Index           =   3
                  Left            =   7890
                  TabIndex        =   59
                  Top             =   330
                  Width           =   555
               End
               Begin VB.CheckBox ChkB1_Youbi 
                  BackColor       =   &H00C0E0FF&
                  Height          =   315
                  Index           =   2
                  Left            =   7320
                  TabIndex        =   58
                  Top             =   330
                  Width           =   555
               End
               Begin VB.CheckBox ChkB1_Youbi 
                  BackColor       =   &H00C0E0FF&
                  Height          =   315
                  Index           =   1
                  Left            =   6750
                  TabIndex        =   57
                  Top             =   330
                  Width           =   555
               End
               Begin VB.CheckBox ChkB1_YoubiHC 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "ÇµÇ»Ç¢"
                  BeginProperty Font 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   5145
                  TabIndex        =   55
                  Top             =   330
                  Width           =   1050
               End
               Begin VB.CheckBox ChkB1_Youbi 
                  BackColor       =   &H00C0E0FF&
                  Height          =   315
                  Index           =   0
                  Left            =   6180
                  TabIndex        =   56
                  Top             =   330
                  Width           =   555
               End
               Begin VB.CheckBox ChkB1_Kara 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "óL"
                  Height          =   315
                  Left            =   4320
                  TabIndex        =   54
                  Top             =   330
                  Width           =   795
               End
               Begin CisText_V60.CisText iB1_UName 
                  Height          =   345
                  Left            =   705
                  TabIndex        =   50
                  Top             =   330
                  Width           =   2595
                  _ExtentX        =   4577
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
                  Text            =   "XXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   20
                  IMEMode         =   4
               End
               Begin CisText_V60.CisText iB1_Ukeir 
                  Height          =   375
                  Left            =   330
                  TabIndex        =   49
                  Top             =   315
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
               Begin CisText_V60.CisText iB1_Cycle1 
                  Height          =   360
                  Left            =   3270
                  TabIndex        =   51
                  Top             =   300
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
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#"
                  cILength        =   1
               End
               Begin CisText_V60.CisText iB1_Cycle2 
                  Height          =   375
                  Left            =   3525
                  TabIndex        =   52
                  Top             =   300
                  Width           =   390
                  _ExtentX        =   688
                  _ExtentY        =   661
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
               Begin CisText_V60.CisText iB1_Cycle3 
                  Height          =   375
                  Left            =   3915
                  TabIndex        =   53
                  Top             =   300
                  Width           =   390
                  _ExtentX        =   688
                  _ExtentY        =   661
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
                  cAutoEject      =   0   'False
               End
            End
         End
         Begin Cis3D_v60.CIS3D Back_Ginko 
            Height          =   1035
            Left            =   60
            Top             =   7365
            Width           =   12300
            _ExtentX        =   21696
            _ExtentY        =   1826
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
            Begin Cis3D_v60.CIS3D UC_3D12 
               Height          =   495
               Index           =   3
               Left            =   6300
               Top             =   510
               Width           =   5970
               _ExtentX        =   10530
               _ExtentY        =   873
               ForeColor       =   16711680
               Caption         =   "ñº  ã`  êl"
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
               cAlingnment     =   1
               cPositionX      =   200
               Begin CisText_V60.CisText iB1_JMeigi 
                  Height          =   375
                  Left            =   1845
                  TabIndex        =   113
                  Top             =   90
                  Width           =   3825
                  _ExtentX        =   6747
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
                  Text            =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
                  MaxLength       =   30
                  IMEMode         =   6
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   495
               Index           =   3
               Left            =   2760
               Top             =   510
               Width           =   3540
               _ExtentX        =   6244
               _ExtentY        =   873
               ForeColor       =   16711680
               Caption         =   "óa ã‡ ãÊ ï™"
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
               cAlingnment     =   1
               cPositionX      =   200
               Begin Cis3D_v60.CIS3D B1lb_JYokin 
                  Height          =   360
                  Left            =   1980
                  Top             =   60
                  Width           =   810
                  _ExtentX        =   1429
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "XXXXXX"
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
               Begin CisText_V60.CisText iB1_JYokin 
                  Height          =   375
                  Left            =   1680
                  TabIndex        =   112
                  Top             =   60
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   661
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
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#"
                  cILength        =   1
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D10 
               Height          =   495
               Index           =   2
               Left            =   2760
               Top             =   15
               Width           =   4815
               _ExtentX        =   8493
               _ExtentY        =   873
               ForeColor       =   16711680
               Caption         =   "ã‚  çs  ñº"
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
               cAlingnment     =   1
               cPositionX      =   200
               Begin CisText_V60.CisText iB1_JGinnm 
                  Height          =   375
                  Left            =   1725
                  TabIndex        =   109
                  Top             =   60
                  Width           =   2565
                  _ExtentX        =   4524
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
                  Text            =   "XXXXXXXXX1XXXXXXXXX2"
                  MaxLength       =   20
                  IMEMode         =   4
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D5 
               Height          =   495
               Index           =   0
               Left            =   330
               Top             =   15
               Width           =   2430
               _ExtentX        =   4286
               _ExtentY        =   873
               ForeColor       =   16711680
               Caption         =   "ã‚çs/éxìX"
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
               cAlingnment     =   1
               cPositionX      =   30
               Begin CisText_V60.CisText iB1_JGinko 
                  Height          =   375
                  Left            =   1185
                  TabIndex        =   107
                  Top             =   60
                  Width           =   600
                  _ExtentX        =   1058
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
                  Text            =   "0000"
                  MaxLength       =   4
                  cDataType       =   1
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "####"
                  cILength        =   4
               End
               Begin CisText_V60.CisText iB1_JShiten 
                  Height          =   375
                  Left            =   1800
                  TabIndex        =   108
                  Top             =   75
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   661
                  cFormat         =   "000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "000"
                  MaxLength       =   3
                  cDataType       =   1
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D12 
               Height          =   495
               Index           =   2
               Left            =   7575
               Top             =   15
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   873
               ForeColor       =   16711680
               Caption         =   "ñ{ éx ìX ñº"
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
               cAlingnment     =   1
               cPositionX      =   200
               Begin CisText_V60.CisText iB1_JHnm 
                  Height          =   375
                  Left            =   1785
                  TabIndex        =   110
                  Top             =   90
                  Width           =   2580
                  _ExtentX        =   4551
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
                  Text            =   "XXXXXXXXX1XXXXXXXXX2"
                  MaxLength       =   20
                  IMEMode         =   4
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   495
               Index           =   2
               Left            =   330
               Top             =   510
               Width           =   2430
               _ExtentX        =   4286
               _ExtentY        =   873
               ForeColor       =   16711680
               Caption         =   " å˚ç¿î‘çÜ"
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
               cAlingnment     =   1
               cPositionX      =   30
               Begin CisText_V60.CisText iB1_JKouza 
                  Height          =   375
                  Left            =   1215
                  TabIndex        =   111
                  Top             =   60
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   661
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
            Begin Cis3D_v60.CIS3D UC_3D16 
               Height          =   990
               Index           =   2
               Left            =   30
               Top             =   15
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   1746
               BackColor       =   33023
               ForeColor       =   0
               Caption         =   "ã‚çsèÓïÒ"
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
               cFont3D         =   2
               cPositionX      =   50
            End
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
      Height          =   975
      Left            =   1440
      ScaleHeight     =   915
      ScaleWidth      =   3330
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   525
      Width           =   3390
      Begin Cis3D_v60.CIS3D B1_Area22 
         Height          =   825
         Left            =   60
         Top             =   60
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1455
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
         cAlingnment     =   6
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   120
            TabIndex        =   0
            Top             =   405
            Width           =   735
            _ExtentX        =   1296
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
            Text            =   "XXXXX"
            MaxLength       =   5
         End
         Begin CisText_V60.CisText iH1_Kojyo 
            Height          =   375
            Left            =   840
            TabIndex        =   1
            Top             =   405
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
         Begin Cis3D_v60.CIS3D l_SortNo 
            Height          =   240
            Index           =   0
            Left            =   0
            Top             =   0
            Width           =   180
            _ExtentX        =   318
            _ExtentY        =   423
            Caption         =   "2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cButton         =   -1  'True
            cPositionX      =   45
            cPositionY      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D35 
         Height          =   825
         Left            =   1365
         Top             =   60
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   1455
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
         Begin Cis3D_v60.CIS3D H1Opb_Torkb 
            Height          =   360
            Left            =   60
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
               Caption         =   "î≠íç"
               Height          =   240
               Index           =   1
               Left            =   900
               TabIndex        =   3
               Top             =   75
               Width           =   825
            End
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
         End
         Begin Cis3D_v60.CIS3D l_SortNo 
            Height          =   240
            Index           =   1
            Left            =   0
            Top             =   0
            Width           =   180
            _ExtentX        =   318
            _ExtentY        =   423
            Caption         =   "1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cButton         =   -1  'True
            cPositionX      =   45
            cPositionY      =   -30
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
      Height          =   795
      Left            =   5250
      ScaleHeight     =   735
      ScaleWidth      =   9525
      TabIndex        =   115
      TabStop         =   0   'False
      Top             =   675
      Width           =   9585
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   6900
         Top             =   75
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
         Left            =   7755
         Top             =   75
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
         Left            =   8610
         Top             =   75
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
         Left            =   3495
         Top             =   75
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
         Left            =   2640
         Top             =   75
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
         Left            =   1785
         Top             =   75
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
         Left            =   75
         Top             =   75
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
         Left            =   930
         Top             =   75
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
         Left            =   6045
         Top             =   75
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
      Begin CisBtn_60.CisBtn PB_CHG 
         Height          =   585
         Left            =   5205
         Top             =   75
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
         BtnCaption      =   "ïœä∑"
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
         PFForeColor     =   16777215
         PFShadowColor   =   0
         PFBackColor     =   0
      End
      Begin CisBtn_60.CisBtn PB_Kotei 
         Height          =   585
         Left            =   4350
         Top             =   75
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
         BtnCaption      =   "çHíˆ"
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
         PFForeColor     =   16777215
         PFShadowColor   =   0
         PFBackColor     =   16711935
      End
   End
   Begin VB.Timer Timer2 
      Left            =   720
      Top             =   420
   End
   Begin VB.Timer Timer1 
      Left            =   300
      Top             =   420
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
      IMEMode         =   3  'µÃå≈íË
      Left            =   0
      TabIndex        =   114
      Text            =   "Text1"
      Top             =   480
      Width           =   255
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
      Caption         =   "Åy éÊà¯êÊÉ}ÉXÉ^ÉÅÉìÉeÉiÉìÉX  Åz"
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
         TabIndex        =   118
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
End
Attribute VB_Name = "CXM0020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :   éÊà¯êÊÉ}ÉXÉ^ÉÅÉìÉeÉiÉìÉX
'**       ÉtÉHÅ[ÉÄID    :   CXM0020
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :   2003/11/17  By CIS SIMPR-A
'**       ïœ  çX  ì˙    :   2004/05/11  By CIS ÉfÅ[É^ëóêMêÊèÓïÒí«â¡
'**       ïœ  çX  ì˙    :   2004/05/25  By CIS çHèÍñºì¸óÕâ¬î\(çHèÍÉRÅ[ÉhÇ»ÇµÇÃéû)
'**                                            çHèÍñºïKê{âèú
'**       ïœ  çX  ì˙    :   2004/06/30  By CIS éÛíçÇÃéûÅAå_ñÒêÊÇÃ¡™Ø∏ŒﬁØ∏ΩÇ™ì¸óÕïsâ¬ÅiíºëOÇÃï\é¶Ç™é–ì‡ÇÃéûÅj
'**       ïœ  çX  ì˙    :   2005/01/31  By CIS î≠íçÇÃéûÅAì`ï[éÌóﬁÇ…é©ããÉ`ÉPÉbÉgÅEéxããÉ`ÉPÉbÉgÇí«â¡
'**       ïœ  çX  ì˙    :   2005/07/15  By CIS ójì˙î≠íçÇÃí«â¡
'**       ïœ  çX  ì˙    :   2006/02/23  By CIS ójì˙î≠íçÇÃï\é¶ÅEâ∫êøãÊï™ÇÃçXêVÇ™Ç≥ÇÍÇ»Ç¢ÅB
'**       ïœ  çX  ì˙    :   2006/03/03  By CIS Key Up/Down éûÇÃÉ\Å[Égèá
'**       ïœ  çX  ì˙    :   2007/12/11  By CIS î[ïièëî≠çsãÊï™•éxããèëî≠çsãÊï™•éÛóÃèëî≠çsãÊï™Å@ÇÃí«â¡
'**                                             ÅEî[ïièëå`ë‘•éÛóÃèëå`ë‘Çï\é¶çÌèú(B1Comb_Nouhin,B1Comb_Jyuryo)
'**                                             Å@ÉfÅ[É^ÇÕé©ããÇÕ1:î≠çsÅAéxããÇÕ2:ÉtÉ@ÉCÉãèoóÕÇ∆Ç∑ÇÈ
'**                                             ÅEñºèÃÉ}ÉXÉ^"ì`ï[éÌóﬁéÛíç•î≠íç"ÇÃâpêîéö1Ç…î[ïièëå`ë‘•éÛóÃèëå`ë‘Çê›íË
'**       ïœ  çX  ì˙    :   2008/02/04  By Cis SIMPR-AïWèÄÇÊÇËà⁄çs
'**                                            çÜï‚ãÊï™ÅAëOçHíˆãÊï™í«â¡
'**       ïœ  çX  ì˙    :   2008/06/11  By Cis åvéZà Çí«â¡
'**                                            Å¶åvéZíPà Ç…3ÅFè⁄ç◊ê›íËÇí«â¡ÅAè⁄ç◊ê›íËéûÇÃÇ›åvéZà ê›íËóL
'**       ïœ  çX  ì˙    :   2008/12/26  By Cis ïœä∑å≥éÊà¯êÊCDí«â¡
'**       ïœ  çX  ì˙    :   2009/01/06  By Cis ïœä∑å≥éÊà¯êÊCDîpé~ÅAéÛíçéÊà¯êÊCDïœä∑œΩ¿í«â¡
'**       ïœ  çX  ì˙    :   2009/07/06  By CIS êVe-Ç©ÇÒÇŒÇÒëŒâû•••î[ì¸êÊéñã∆ëÃ/édì¸êÊéñã∆ëÃ/èoâ◊èÍ í«â¡
'**       ïœ  çX  ì˙    :   2009/07/15  By CIS ìyäÚí√ópÉJÉXÉ^É}ÉCÉY
'**       ïœ  çX  ì˙    :   2009/11/06  By CIS ïiî‘É`ÉFÉbÉNåÖêîí«â¡
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' Ã´∞—±∏√®Ãﬁ Ã◊∏ﬁ
    Dim ProcHB          As String * 2   ' ì¸óÕàÊêÿë÷
    Dim SyoriKB         As String       ' èàóùãÊï™(êVãK•èCê≥•çÌèú ìô)
    Dim SyoriNM         As String       ' èàóùñºèÃ(êVãK•èCê≥•çÌèú ìô)
    Dim wChk            As Boolean

    Dim mGotColor       As Long
    Dim mLostColor      As Long
    
    Dim mIKeisan        As Integer      ' åvéZï˚ñ@èâä˙íl
    Dim mIMarume        As Integer      ' åvéZíPà èâä˙íl
    '------------------------------------------------------------***** 2008/06/11 Start
    Dim mIKurai         As Integer      ' åvéZà èâä˙íl
    '------------------------------------------------------------***** 2008/06/11 End
    Dim mIHei           As Integer      ' ïΩèÄâªèâä˙íl

'''2007/12/11DEL    Dim CisComb1        As CisCombo
    Dim CisComb2        As CisCombo
'''2007/12/11DEL    Dim CisComb3        As CisCombo
    Dim CisComb4        As CisCombo

    
    Dim CisVsGrid       As New CisVsGrid3
    
    Dim ActObj          As Object
    Dim ActRow          As Integer

                                        ' < 2006/03/03 ADD >
    Dim udKey           As Integer      ' 1:éÊà¯êÊCD•éÊà¯êÊãÊï™ÅA0:éÊà¯êÊãÊï™•éÊà¯êÊCD
'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2007/12/11 Update Satart
    Dim SaveKeitai      As String
'Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™Ñ™ 2007/12/11 Update End

'============================================ 2009/01/06 Start
    Dim ChgTorcd_Flg    As Boolean
    Dim CisVsGrid1      As New CisVsGrid3
'============================================ 2009/01/06 End

'======================== 2009/07/15 Start
    Dim CisComb5        As CisCombo
    Dim CisComb6        As CisCombo
    Dim CisVsGrid2      As New CisVsGrid3
    Dim ChgKotei_Flg    As Boolean
'======================== 2009/07/15 End
    

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
    
    mGotColor = B1Op_Keisan(0).BackColor
    mLostColor = B1Op_Keisan(1).BackColor
    B1Op_Keisan(0).BackColor = mLostColor
    
    Dummy.Left = -1000
    Dummy.Enabled = False
    Timer1.Interval = 1000   ' éûçèï\é¶ä‘äu
    Timer1.Enabled = True    ' éûçèï\é¶äJén
    Timer2.Interval = 500    ' èàóùãÊï™ì_ñ≈ä‘äu
    Timer2.Enabled = False   ' èàóùãÊï™ì_ñ≈ã÷é~
    
    mIKeisan = 0
    mIMarume = 0
    mIHei = 0
    '-------------------------------------------------------------***** 2008/06/11 Start
    mIKurai = 0
    '-------------------------------------------------------------***** 2008/06/11 End
    
    ' ñºèÃÉ}ÉXÉ^ämîF
    Call MeisyoInit
'    '+---------------------+
'    '+ ∫›ŒﬁŒﬁØ∏Ω èâä˙ê›íË  +
'    '+---------------------+
'''2007/12/11DEL    Set CisComb1 = New CisCombo
'''2007/12/11DEL    With CisComb1
'''2007/12/11DEL        Set .Connect = CisDB
'''2007/12/11DEL        Set .CombObj = B1Comb_Jyuryo
'''2007/12/11DEL        .Control = "ñºèÃÉ}ÉXÉ^"
'''2007/12/11DEL        .ComboName = "éÛóÃèëå`ë‘"
'''2007/12/11DEL        .Ess = False
'''2007/12/11DEL        .IniPath = gIniExe
'''2007/12/11DEL        .SetSpace = 50
'''2007/12/11DEL        .ComboInit
'''2007/12/11DEL    End With
    Set CisComb2 = New CisCombo
    With CisComb2
        Set .Connect = CisDB
        Set .CombObj = B1Comb_HNaiji
        .Control = "ñºèÃÉ}ÉXÉ^"
        .ComboName = "î≠íçì‡é¶"
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .ComboInit
    End With
''    Set CisComb3 = New CisCombo
''    With CisComb3
''        Set .Connect = CisDB
''        Set .CombObj = B1Comb_Nouhin
''        .Control = "ñºèÃÉ}ÉXÉ^"
''        .ComboName = "î[ïièëå`ë‘éÛíç"
''        .Ess = False
''        .IniPath = gIniExe
''        .SetSpace = 50
''        .ComboInit
''    End With
    Set CisComb4 = New CisCombo
    With CisComb4
        Set .Connect = CisDB
        Set .CombObj = B1Comb_Meaisai
        .Control = "ñºèÃÉ}ÉXÉ^"
        .ComboName = "ñæç◊èëî≠çs"
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .ComboInit
    End With
    
    '============================ 2009/07/15 Start
    Set CisComb5 = New CisCombo
    With CisComb5
        Set .Connect = CisDB
        Set .CombObj = B1Comb_Bumon1
        .Control = "ñºèÃÉ}ÉXÉ^"
        .ComboName = "ïîñÂ1"
        .Ess = False
        .IniPath = gIniExe
        .SetSpace = 50
        .ComboInit
    End With
    Set CisComb6 = New CisCombo
    With CisComb6
        Set .Connect = CisDB
        Set .CombObj = B1Comb_Bumon2
        .Control = "ñºèÃÉ}ÉXÉ^"
        .ComboName = "ïîñÂ2"
        .Ess = False
        .IniPath = gIniExe
        .SetSpace = 50
        .ComboInit
    End With
    '============================ 2009/07/15 End
    
    
    '+-----------------------+
    '+  ÉOÉäÉbÉhèâä˙ê›íË
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = vsGrid
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 19
        .InitGet = False
        .Init
    End With
    Call GridInit
'   #------------------#
'   #    èâä˙ì¸óÕê›íË   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
    Call H1Op_Torkb_Click(0)        ' 2009.11.06 add
'   #------------------#
'   #  èàóùãÊï™èâä˙ï\é¶ #
'   #------------------#
    SyoriKB = "ADD"
    Call SyoriName(SyoriKB)
'   #-------------------#
'   # âÊñ íÜâõà íuï\é¶   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    '--------------------------------------------------***** 2008/06/09 Start
    Back_Kurai.Visible = False
    '--------------------------------------------------***** 2008/06/09 End
    
    '============================================= 2008/12/26 Start
    Back_ConvTorcd1.Visible = False
    Back_ConvTorcd2.Visible = True
    '============================================= 2008/12/26 End
    
    '============================================= 2009/01/06 Start
    Back_ChgTorcd.Visible = False
    ChgTorcd_Flg = False
    
    '+------------------------------------+
    '+  éÊà¯êÊ∫∞ƒﬁïœä∑ì¸óÕÉOÉäÉbÉhèâä˙ê›íË
    '+------------------------------------+
    With CisVsGrid1
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 5
        .InitGet = False
        .Init
    End With
    Call GridInit1
    '============================================= 2009/01/06 End
    
    
    '============================================= 2009/01/06 Start
    Back_Kotei.Visible = False
    ChgKotei_Flg = False
    
    '+------------------------------------+
    '+  çHíˆèÓïÒì¸óÕÉOÉäÉbÉhèâä˙ê›íË
    '+------------------------------------+
    With CisVsGrid2
        Set .GridObj = VsGrid2
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 12
        .InitGet = False
        .Init
    End With
    Call GridInit2
    '============================================= 2009/01/06 End
    
    
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    ÅsÇ`ÇÉÇîÇâÇñÇÅÇîÇÖ ÉCÉxÉìÉgÅt                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    H1Op_Torkb(1).Value = True
    Call H1Op_Torkb_Click(1)
    B1Chk_Juryo.Value = 1
    
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
                If ProcHB = "B1" Then
                   If iB1_Ukeir.Visible Then
                      iB1_Ukeir.SetFocus
                   End If
                End If
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    'ÅyèCê≥Åz
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    'ÅyçÌèúÅz
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  'ÅyåüçıÅz
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  'Åyå„ëﬁÅz
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  'ÅyëOêiÅz
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    'ÅyéÊè¡Åz
'======================================================================= 2009/07/15 Start
'           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    'ÅyèIóπÅz
           Case vbKeyF12
                               If PB_END.Visible Then
                                   Call PB_END_Click    'ÅyèIóπÅz
                               Else
                                  If PB_KoteiEND.Visible Then
                                     Call PB_KoteiEND_Click    'ÅyçHíˆèÓïÒì¸óÕèIóπÅz
                                  End If
                               End If
'======================================================================= 2009/07/15 End
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      'Åyì¸óÕÅz
                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
                                  Dummy.Enabled = False
                               End If
'-----     -----     -----     -----     -----     -----     -----     < 2006/03/03 ADD START >
            Case vbKeyS
                                If ProcHB = "H1" And Shift = 1 Then
                                    Call l_SortNo_Click(0)
                                Else
                                    Exit Function
                                End If
'-----     -----     -----     -----     -----     -----     -----     < 2006/03/03 ADD  END  >
'======================================================================= 2009/01/06 Start
           Case vbKeyF6:       If PB_CHG.Visible Then Call PB_CHG_Click    'Åyïœä∑ÅzéÊà¯êÊ∫∞ƒﬁïœä∑ì¸óÕ
'======================================================================= 2009/01/06 End
'======================================================================= 2009/07/15 Start
           Case vbKeyF5:       If PB_Kotei.Visible Then Call PB_Kotei_Click    'ÅyçHíˆÅzçHíˆèÓïÒì¸óÕ
'======================================================================= 2009/07/15 End
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
    If SyoriKB <> "ADD" Then
        SyoriKB = "ADD"
        Call SyoriName(SyoriKB)
    End If
    iH1_Torcd.SetFocus
End Sub
'+----------------------------+
'+    èC  ê≥  ÉL Å[(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If SyoriKB <> "MNT" Then
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
    End If
    iH1_Torcd.SetFocus
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
    iH1_Torcd.SetFocus
End Sub
'=============================================== 2009/01/06 Start
'+----------------------------+
'+    ïœÅ@ä∑  ÉL Å[(F6)       +
'+----------------------------+
Private Sub PB_CHG_Click()
    If Not ChgTorcd_Flg Then
       ' éÛíçéÊà¯êÊCDïœä∑É}ÉXÉ^èÓïÒ
       gSL_Select = "Select * from éÛíçéÊà¯êÊCDïœä∑É}ÉXÉ^ "
       gSL_Select = gSL_Select & " Where ïœä∑êÊCD = '" & RTrim(iH1_Torcd) & RTrim(iH1_Kojyo) & "'"
       gSL_Select = gSL_Select & " Order By ïœä∑å≥CD,ïœä∑å≥éÛì¸ "
       If JTHRead(gSL_Select, 0, 1) Then
          Call GridSet1
       End If
       Call JTHClose(1)
       
       vsGrid1.Row = 1
       Call VsGridToText1(1)
       
       Call TextToVsGrid(Val(iB1_Ukeir.Tag))
       
       Back_ChgTorcd.Visible = True
       ChgTorcd_Flg = True
       PB_CAN.Visible = False
       PB_ENT.Visible = False
       PB_CHG.BtnCaption = "èIóπ"
       B1lb_Migiya = "Å®"
       
       iB1_MotoCD.SetFocus
       B1_Area1.Enabled = False
       
    Else
       Call TextToVsGrid1(Val(iB1_MotoCD.Tag))
       If Not Grid1Chk Then Exit Sub            'ì¸óÕì‡óe¡™Ø∏
       Call VsGrid1_DBInsert                    'ìoò^ì‡óeDBçXêV

       Back_ChgTorcd.Visible = False
       ChgTorcd_Flg = False
       PB_CAN.Visible = True
       PB_ENT.Visible = True
       PB_CHG.BtnCaption = "ïœä∑"
    
       CisVsGrid1.Clear
       B1_Area1.Enabled = True
    End If
End Sub
'=============================================== 2009/01/06 End
'=============================================== 2009/07/15 Start
'+----------------------------+
'+    çHíˆâÊñ èIóπ  ÉL Å[     +
'+----------------------------+
Private Sub PB_KoteiEND_Click()
    Back_Kotei.Visible = False
    PNL_PF.Visible = True
    B1_Area1.Enabled = True
    iB1_Tornm.SetFocus
End Sub
'+----------------------------+
'+    çHíˆâÊñ ï\é¶  ÉL Å[     +
'+----------------------------+
Private Sub PB_Kotei_Click()
    Call VsGridToText2(1)
    Back_Kotei.Visible = True
    B1Comb_Bumon1.SetFocus
    PNL_PF.Visible = False
    B1_Area1.Enabled = False
End Sub
'=============================================== 2009/07/15 End
'+----------------------------+
'+    å„  ëﬁ  ÉL Å[(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
'-----     -----     -----     -----     -----     -----     -----     -----     < 2006/03/03 MNT START >
''*--------------------------------------------------<< 2004/06/30 Update Start >>
''   gSL_Select = " SELECT *"
''   gSL_Select = gSL_Select & " FROM éÊà¯êÊÉ}ÉXÉ^ "
''   gSL_Select = gSL_Select & " WHERE éÊà¯êÊCD + Cast(éÊà¯êÊãÊï™ as Varchar(1) ) "
''   gSL_Select = gSL_Select & "   < '" & RTrim(iH1_Torcd) & RTrim(iH1_Kojyo) & Format(H1Op_Torkb(0).Tag, "0") & "'"
''   gSL_Select = gSL_Select & " ORDER BY éÊà¯êÊCD + Cast(éÊà¯êÊãÊï™ as Varchar(1) ) Desc "
'    gSL_Select = "select * from éÊà¯êÊÉ}ÉXÉ^"
'    gSL_Select = gSL_Select & " where éÊà¯êÊCD + space(7 - len(éÊà¯êÊCD))  + cast(éÊà¯êÊãÊï™ as varchar(1))"
'    gSL_Select = gSL_Select & " < "
'    gStr = RTrim(iH1_Torcd) & RTrim(iH1_Kojyo)
'    gSL_Select = gSL_Select & "'" & gStr & Space(7 - CisFun.Len2(gStr)) & Format(H1Op_Torkb(0).Tag, "0") & "'"
'    gSL_Select = gSL_Select & " order by éÊà¯êÊCD desc, éÊà¯êÊãÊï™ desc"
''*--------------------------------------------------<< 2004/06/30 Update End >>
    gSL_Select = "select * from éÊà¯êÊÉ}ÉXÉ^"
    
    If udKey = 1 Then
        gSL_Select = gSL_Select & " where éÊà¯êÊCD + space(7 - len(éÊà¯êÊCD))  + cast(éÊà¯êÊãÊï™ as varchar(1))"
        gSL_Select = gSL_Select & " < "
        gStr = RTrim(iH1_Torcd) & RTrim(iH1_Kojyo)
        gSL_Select = gSL_Select & "'" & gStr & Space(7 - CisFun.Len2(gStr)) & Format(H1Op_Torkb(0).Tag, "0") & "'"
        gSL_Select = gSL_Select & " order by éÊà¯êÊCD desc, éÊà¯êÊãÊï™ desc"
    Else
        gSL_Select = gSL_Select & " where cast(éÊà¯êÊãÊï™ as varchar(1)) + éÊà¯êÊCD + space(7 - len(éÊà¯êÊCD))"
        gSL_Select = gSL_Select & " < "
        gStr = RTrim(iH1_Torcd) & RTrim(iH1_Kojyo)
        gSL_Select = gSL_Select & "'" & Format(H1Op_Torkb(0).Tag, "0") & gStr & Space(7 - CisFun.Len2(gStr)) & "'"
        gSL_Select = gSL_Select & " order by éÊà¯êÊãÊï™ desc, éÊà¯êÊCD desc"
    End If
'-----     -----     -----     -----     -----     -----     -----     -----     < 2006/03/03 MNT  END  >
    
    Call DBInput("Back")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Torcd.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    ëO  êi  ÉL Å[(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    
    Call CisFrm.MousePT(11)
    
'-----     -----     -----     -----     -----     -----     -----     -----     < 2006/03/03 MNT START >
''*--------------------------------------------------<< 2004/06/30 Update Start >>
''   gSL_Select = " SELECT *"
''   gSL_Select = gSL_Select & " FROM éÊà¯êÊÉ}ÉXÉ^ "
''   gSL_Select = gSL_Select & " WHERE éÊà¯êÊCD "
''   gSL_Select = gSL_Select & "   > '" & RTrim(iH1_Torcd) & RTrim(iH1_Kojyo) & "'"
''   gSL_Select = gSL_Select & "    And éÊà¯êÊãÊï™ = " & CisFun.Val2(H1Op_Torkb(0).Tag)
''   gSL_Select = gSL_Select & " ORDER BY éÊà¯êÊCD "
'    gSL_Select = "select * from éÊà¯êÊÉ}ÉXÉ^"
'    gSL_Select = gSL_Select & " where éÊà¯êÊCD + space(7 - len(éÊà¯êÊCD))  + cast(éÊà¯êÊãÊï™ as varchar(1))"
'    gSL_Select = gSL_Select & " > "
'    gStr = RTrim(iH1_Torcd) & RTrim(iH1_Kojyo)
'    gSL_Select = gSL_Select & "'" & gStr & Space(7 - CisFun.Len2(gStr)) & Format(H1Op_Torkb(0).Tag, "0") & "'"
'    gSL_Select = gSL_Select & " order by éÊà¯êÊCD, éÊà¯êÊãÊï™"
''*--------------------------------------------------<< 2004/06/30 Update End >>
    gSL_Select = "select * from éÊà¯êÊÉ}ÉXÉ^"
    If udKey = 1 Then
        gSL_Select = gSL_Select & " where éÊà¯êÊCD + space(7 - len(éÊà¯êÊCD))  + cast(éÊà¯êÊãÊï™ as varchar(1))"
        gSL_Select = gSL_Select & " > "
        gStr = RTrim(iH1_Torcd) & RTrim(iH1_Kojyo)
        gSL_Select = gSL_Select & "'" & gStr & Space(7 - CisFun.Len2(gStr)) & Format(H1Op_Torkb(0).Tag, "0") & "'"
        gSL_Select = gSL_Select & " order by éÊà¯êÊCD, éÊà¯êÊãÊï™"
    Else
        gSL_Select = gSL_Select & " where cast(éÊà¯êÊãÊï™ as varchar(1)) + éÊà¯êÊCD + space(7 - len(éÊà¯êÊCD))"
        gSL_Select = gSL_Select & " > "
        gStr = RTrim(iH1_Torcd) & RTrim(iH1_Kojyo)
        gSL_Select = gSL_Select & "'" & Format(H1Op_Torkb(0).Tag, "0") & gStr & Space(7 - CisFun.Len2(gStr)) & "'"
        gSL_Select = gSL_Select & " order by éÊà¯êÊãÊï™, éÊà¯êÊCD"
    End If
'-----     -----     -----     -----     -----     -----     -----     -----     < 2006/03/03 MNT  END  >
    
    Call DBInput("Next")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Torcd.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'-----     -----     -----     -----     -----     < 2006/03/03 MNT START >
'+-----------------------------+
'+      UP/Down ÉLÅ[èá
'+-----------------------------+
Private Sub l_SortNo_Click(INDEX As Integer)
    If udKey = 0 Then
        udKey = 1
        l_SortNo(0) = "1"
        l_SortNo(1) = "2"
    Else
        udKey = 0
        l_SortNo(0) = "2"
        l_SortNo(1) = "1"
    End If
End Sub
'-----     -----     -----     -----     -----     < 2006/03/03 MNT  END  >

'+----------------------------+
'+    åü  çı  ÉL Å[(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Kojyo" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iB1_KTorcd" Then Call Look_KTorcd: Exit Sub
    If PB_Look.Tag = "iB1_DenSyu" Then Call Look_Kubun(1): Exit Sub
    If PB_Look.Tag = "iB1_JYokin" Then Call Look_Kubun(0): Exit Sub
    If PB_Look.Tag = "iB1_STHYmd" Then iB1_STHYmd.ShowCalender: Exit Sub
    If Mid(PB_Look.Tag, 3) = "iB1_SendFax" Then Call Look_STorcd("F", CisFun.Val2(Mid(PB_Look.Tag, 1, 2))): Exit Sub
    If Mid(PB_Look.Tag, 3) = "iB1_SendMail" Then Call Look_STorcd("M", CisFun.Val2(Mid(PB_Look.Tag, 1, 2))): Exit Sub
End Sub
'+----------------------------+
'+       éÊà¯êÊåüçı            +
'+----------------------------+
Private Sub Look_Torcd()

    RV_Left = 0
    RV_Top = 0
    RV_TorKb = CisFun.Val2(H1Op_Torkb(0).Tag)
    
    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

    If RV_Rtn Then
       iH1_Torcd = RV_Torcd
       iH1_Kojyo = RV_Kojyo
    End If

End Sub
'+----------------------------+
'+       édì¸êÊåüçı            +
'+----------------------------+
Private Sub Look_STorcd(CType As String, KB As Integer)

    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 8
    RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
       If CType = "F" Then
          iB1_SendFax(KB) = RV_TorcdK
          Call iB1_SendFax_LostFocus(KB)
       Else
          iB1_SendMail(KB) = RV_TorcdK
          Call iB1_SendMail_LostFocus(KB)
       End If
    End If
End Sub
'+----------------------------+
'+       å_ñÒêÊåüçı(ìæà”êÊ)    +
'+----------------------------+
Private Sub Look_KTorcd()
    RV_Left = 0
    RV_Top = 0
    If CisFun.Val2(H1Op_Torkb(0).Tag) = 0 Then
        RV_TorKb = 2
    Else
        RV_TorKb = 5
    End If
    
    RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iB1_KTorcd = RV_TorcdK
        B1lb_KTorNm = RV_TorRName
    End If
End Sub
'+----------------------------+
'+       ñºèÃÉ}ÉXÉ^åüçı        +
'+    KB : 0     óaã‡ãÊï™     +
'+       : 0à»äO ì`ï[éÌóﬁ      +
'+----------------------------+
Private Sub Look_Kubun(KB As Integer)

    RV_Left = 0
    RV_Top = 0

    If KB = 1 Then
        If H1Op_Torkb(1).Value Then
            KB = 2
        End If
    End If
    Select Case KB
        Case 0: RVI_Kubun = "óaã‡ãÊï™"
        Case 1: RVI_Kubun = "ì`ï[éÌóﬁéÛíç"
        Case Else: RVI_Kubun = "ì`ï[éÌóﬁî≠íç"
    End Select
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
        Select Case KB
            Case 0
                iB1_JYokin = RV_Code
                Call iB1_JYokin_LostFocus
            Case Else
                iB1_DenSyu = RV_Code
                Call iB1_DenSyu_LostFocus
        End Select
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
End Sub
'+---------------------+
'+   H1Op_Torkb
'+---------------------+
Private Sub H1Op_Torkb_GotFocus(INDEX As Integer)
'    H1Op_Torkb(0).Tag = Index
    H1Op_Torkb(0).BackColor = gIGotSel
    H1Op_Torkb(1).BackColor = gIGotSel
    H1Opb_Torkb.BackColor = gIGotSel
    If INDEX = 0 Then
       B1Chk_Torkb(1).Value = 0
       B1Chk_Torkb(1).Enabled = False
       B1Chk_Juryo.Caption = "èoâ◊éû"
    Else
       B1Chk_Torkb(1).Enabled = True
       If B1Chk_Torkb(0).Value = 1 Then
          B1Chk_Torkb(1).Value = 0
          B1Chk_Torkb(1).Enabled = False
       Else
          B1Chk_Torkb(1).Enabled = True
       End If
       B1Chk_Juryo.Caption = "éÛóÃéû"
    End If
    B1Chk_Juryo.Value = 1
End Sub
Private Sub H1Op_Torkb_LostFocus(INDEX As Integer)
    H1Op_Torkb(0).BackColor = gILostSel
    H1Op_Torkb(1).BackColor = gILostSel
    H1Opb_Torkb.BackColor = gILostSel
End Sub
Private Sub H1Op_Torkb_Click(INDEX As Integer)
    H1Op_Torkb(0).Tag = Format(INDEX, "0")
    '---------------------------------------' 2009.11.06 add start
    iB1_HinKeta = ""
    If INDEX = 0 Then
       Back_HinKeta.Visible = True
    Else
       Back_HinKeta.Visible = False
    '---------------------------------------' 2009.11.06 add end
    End If
End Sub
'+-----------------------------+
'+  iH1_Kojyo
'+-----------------------------+
Private Sub iH1_Kojyo_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Kojyo_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  B1Chk_TorkbÅ@éÊà¯êÊãÊï™
'+-----------------------------+
Private Sub B1Chk_Torkb_Click(INDEX As Integer)
    Dim wSelectFlg  As Boolean

    wSelectFlg = False
    For Each gObj In B1Chk_Torkb
        If gObj.Value = 1 Then
            wSelectFlg = True
            gObj.BackColor = mGotColor
        Else
            gObj.BackColor = mLostColor
        End If
    Next gObj
    B1Chk_Torkb(0).Tag = INDEX
    
    Call BodyEnbled
    If INDEX = 1 Then
       If iB1_Ukeir.Visible Then
          Call GridTextEnabled(True)
       End If
       With vsGrid
            For gCnt1 = 1 To vsGrid.Rows - 1
                For gCnt2 = 3 To .Cols - 1
                    .TextMatrix(gCnt1, gCnt2) = ""
                Next gCnt2
            Next gCnt1
       End With
    End If

    '=============================== *** 2009/07/06 Start
    If H1Op_Torkb(0) Then
       If B1Chk_Torkb(0).Value = 1 Then
          Back_Maekotei(0).Visible = False
          Back_Maekotei(1).Visible = False
          Back_NJigyo.Visible = True
          Back_SJigyo.Visible = True
          Back_Syukaba.Visible = True
       Else
          Back_Maekotei(0).Visible = False
          Back_Maekotei(1).Visible = True
          Back_NJigyo.Visible = True
          Back_SJigyo.Visible = False
          Back_Syukaba.Visible = True
       End If
    End If
    '=============================== *** 2009/07/06 End

    '=============================== 2009/07/15 Start
    If B1Chk_Torkb(1).Value = 1 Then
       PB_Kotei.Visible = True
    Else
       PB_Kotei.Visible = False
    End If
    '=============================== 2009/07/15 End
    
End Sub
'+--------------------------------------+
'+      ì¸óÕì‡óeÇÃêÿë÷ÇçsÇ§
'+--------------------------------------+
Private Sub BodyEnbled()
    Dim JKeiFlg         As Boolean
    Dim JKeiNoFlg       As Boolean
    Dim HKeiFlg         As Boolean
    Dim HKeiNoFlg       As Boolean
    Dim HSyaFlg         As Boolean
    
    On Error GoTo BodyEnbled_Er
    
    JKeiFlg = False
    JKeiNoFlg = False
    HKeiFlg = False
    HKeiNoFlg = False
    HSyaFlg = False
    
    gInt = CisFun.Val2(H1Op_Torkb(0).Tag)
        
    If gInt = 0 Then
'*--------------------------------------------------<< 2004/06/30 Insert Start >>
       B1Chk_Torkb(0).Enabled = True
'*--------------------------------------------------<< 2004/06/30 Insert End >>
       B1Chk_Torkb(1).Value = 0
       B1Chk_Torkb(1).Enabled = False
    Else
       B1Chk_Torkb(0).Enabled = True
       B1Chk_Torkb(1).Enabled = True
       If B1Chk_Torkb(0).Value = 1 Then
          B1Chk_Torkb(1).Value = 0
          B1Chk_Torkb(1).Enabled = False
       End If
       If B1Chk_Torkb(1).Value = 1 Then
          B1Chk_Torkb(0).Value = 0
          B1Chk_Torkb(0).Enabled = False
          B1Chk_Torkb(1).Enabled = True
       End If
    End If
    
    If gInt = 0 Then
       If B1Chk_Torkb(0).Value = 1 Then
          JKeiFlg = True
       Else
          JKeiNoFlg = True
       End If
    Else
       If B1Chk_Torkb(0).Value = 1 Then
          HKeiFlg = True
       Else
          HKeiNoFlg = True
       End If
       If B1Chk_Torkb(1).Value = 1 Then
          HKeiFlg = False
          HKeiNoFlg = False
          HSyaFlg = True
       End If
    End If
'====================================== 2008/12/26 Start
    If H1Op_Torkb(0) Then
       'Back_ConvTorcd1.Visible = True
       'Back_ConvTorcd2.Visible = False
       Back_ConvTorcd1.Visible = False
       Back_ConvTorcd2.Visible = True
    Else
       Back_ConvTorcd1.Visible = False
       Back_ConvTorcd2.Visible = True
    End If
'====================================== 2008/12/26 End
    
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
    Back_MK.Visible = Not (gInt = 0)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
    
    ' îÌé©é–CD
    iB1_HTorcd.Visible = Not HSyaFlg
    iB1_HKojyo.Visible = Not HSyaFlg
    ' å_ñÒêÊ
    iB1_KTorcd.Visible = JKeiNoFlg Or HKeiNoFlg
    B1lb_KTorNm.Visible = JKeiNoFlg Or HKeiNoFlg
'    'çHèÍñº
'    If Trim(iH1_Kojyo) = "" Then
'       iB1_KojyoNm.Visible = False
'    Else
'       iB1_KojyoNm.Visible = True
'    End If
    ' â∫êøñ@ëŒè€
    B1Chk_Taisyo.Visible = HKeiFlg Or HKeiNoFlg
'----------------------------------------------------------- 2006/02/23 Delete
'    If Not B1Chk_Taisyo Then
'----------------------------------------------------------- 2006/02/23 Delete End
'----------------------------------------------------------- 2006/02/23 Insert
    If Not B1Chk_Taisyo.Visible Then
'----------------------------------------------------------- 2006/02/23 Insert End
       B1Chk_Taisyo.Value = 0
    End If
    ' í˜ì˙
    iB1_Shime.Visible = JKeiFlg Or HKeiFlg
    ' åvéZï˚ñ@
    Back_KeisanT.Visible = JKeiFlg Or HKeiFlg
    Back_KeisanH.Visible = JKeiFlg Or HKeiFlg
    ' ïΩèÄâª
    Back_Heijyun.Visible = HKeiFlg Or HKeiNoFlg
    iB1_Keisu.Visible = HKeiFlg Or HKeiNoFlg
    ' ì`ï[éÌóﬁ
    iB1_DenSyu.Visible = JKeiFlg Or JKeiNoFlg Or HKeiFlg Or HKeiNoFlg
    B1lb_DenSyu.Visible = JKeiFlg Or JKeiNoFlg Or HKeiFlg Or HKeiNoFlg
'''2007/12/11DEL    ' î[ïièëå`ë‘
'''2007/12/11DEL    B1Comb_Nouhin.Visible = JKeiFlg Or JKeiNoFlg Or HKeiFlg Or HKeiNoFlg
    
    ' î[ïièëî≠çsãÊï™
    B1Chk_PrtkbN.Visible = JKeiFlg Or JKeiNoFlg Or HKeiFlg Or HKeiNoFlg
    ' éxããèëî≠çsãÊï™
    B1Chk_PrtkbS.Visible = JKeiFlg Or JKeiNoFlg Or HKeiFlg Or HKeiNoFlg
    ' éÛóÃèëî≠çsãÊï™
    B1Chk_PrtkbJ.Visible = JKeiFlg Or JKeiNoFlg Or HKeiFlg Or HKeiNoFlg
    
'''2007/12/11DEL    ' éÛóÃèëå`ë‘
'''2007/12/11DEL    B1Comb_Jyuryo.Visible = JKeiFlg Or JKeiNoFlg Or HKeiFlg Or HKeiNoFlg
    ' éÛóÃèëî≠çs
    B1Chk_Juryo.Visible = JKeiFlg Or JKeiNoFlg Or HKeiFlg Or HKeiNoFlg
    ' ñæç◊èëî≠çs
    B1Comb_Meaisai.Visible = HKeiFlg Or HKeiNoFlg
    ' î≠íçì‡é¶
    B1Comb_HNaiji.Visible = HKeiFlg Or HKeiNoFlg Or HSyaFlg
    ' ëOì|ì˙êî
    iB1_Mae.Visible = HKeiFlg Or HKeiNoFlg
    ' ã‚çsèÓïÒ
'    Back_Ginko.Visible = JKeiFlg Or HKeiFlg
    Back_Ginko.Visible = False          ' 2004/05/11
    'î≠íçèÓïÒ
    Back_HC.Visible = HKeiFlg Or HKeiNoFlg
    '√ﬁ∞¿ëóêMèÓïÒ
    Back_Send.Visible = HKeiFlg Or HKeiNoFlg
    
    'éwé¶èëî≠çsãÊï™
    B1Chk_Sijisyo.Visible = JKeiFlg Or JKeiNoFlg            '2009/07/15
    
    Exit Sub
BodyEnbled_Er:
    Call CisFun.ErrorBox
End Sub
'+-----------------------------+
'+  ó™èÃ
'+-----------------------------+
Private Sub iB1_Ryaku_LostFocus()
    B1lb_Ryaku1 = CisFun.Mid2(iB1_Ryaku, 1, 10)
    B1lb_Ryaku2 = CisFun.Mid2(iB1_Ryaku, 11)
End Sub
'+-----------------------------+
'+  å_ñÒêÊ
'+-----------------------------+
Private Sub iB1_KTorcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_KTorcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    B1lb_KTorNm = ""

    If CisFun.Val2(H1Op_Torkb(0).Tag) = 0 Then
        If TorNmGet(iB1_KTorcd, 2) Then
            B1lb_KTorNm = TRM_RName
        End If
    Else
        If TorNmGet(iB1_KTorcd, 5) Then
            B1lb_KTorNm = TRM_RName
        End If
    End If
End Sub

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
'+------------------------------+
'+     B1Chk_Gouho çÜï‚ãÊï™     +
'+------------------------------+
Private Sub B1Chk_Gouho_Click()
    If B1Chk_Gouho.Value = 0 Then
       B1Chk_Gouho.Caption = "çÜå˚"
    Else
       B1Chk_Gouho.Caption = "ï‚ãã"
    End If
End Sub
'+---------------------------------+
'+     B1Chk_MaeHak ëOî≠çsãÊï™     +
'+---------------------------------+
Private Sub B1Chk_MaeHak_Click()
    If B1Chk_MaeHak.Value = 0 Then
       B1Chk_MaeHak.Caption = "ï\é¶ñ≥"
    Else
       B1Chk_MaeHak.Caption = "ï\é¶óL"
    End If
End Sub
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END

'+-----------------------------+
'+  B1Chk_Taisyo â∫êøñ@
'+-----------------------------+
Private Sub B1Chk_Taisyo_Click()
    If B1Chk_Taisyo.Value = 0 Then
       B1Chk_Taisyo.Caption = "ëŒè€äO"
    Else
       B1Chk_Taisyo.Caption = "ëŒè€"
    End If
End Sub
'+-----------------------------+
'+  B1Op_Keisan åvéZï˚ñ@
'+-----------------------------+
Private Sub B1Op_Keisan_Click(INDEX As Integer)
    
    B1Op_Keisan(0).Tag = INDEX
    For Each gObj In B1Op_Keisan
        If gObj.INDEX = INDEX Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
'+-----------------------------+
'+  B1Op_Marume åvéZíPà 
'+-----------------------------+
Private Sub B1Op_Marume_Click(INDEX As Integer)
    B1Op_Marume(0).Tag = INDEX
    For Each gObj In B1Op_Marume
        If gObj.INDEX = INDEX Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj

'----------------------------------------------------------------***** 2008/06/11 Start
    If INDEX = 3 Then
       Back_Kurai.Visible = True
    Else
       Back_Kurai.Visible = False
    End If
'----------------------------------------------------------------***** 2008/06/11 End

End Sub
'----------------------------------------------------------------***** 2008/06/11 Start
'+-----------------------------+
'+  B1Op_Kurai1 åvéZà 1åÖñ⁄
'+-----------------------------+
Private Sub B1Op_Kurai1_Click(INDEX As Integer)
    B1Op_Kurai1(0).Tag = INDEX
    For Each gObj In B1Op_Kurai1
        If gObj.INDEX = INDEX Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
'+-----------------------------+
'+  B1Op_Kurai2 åvéZà 2åÖñ⁄
'+-----------------------------+
Private Sub B1Op_Kurai2_Click(INDEX As Integer)
    B1Op_Kurai2(0).Tag = INDEX
    For Each gObj In B1Op_Kurai2
        If gObj.INDEX = INDEX Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
'----------------------------------------------------------------***** 2008/06/11 End
'+-----------------------------+
'+  B1Op_Hei ïΩèÄâªãÊï™
'+-----------------------------+
Private Sub B1Op_Hei_Click(INDEX As Integer)
    B1Op_Hei(0).Tag = INDEX
    For Each gObj In B1Op_Hei
        If gObj.INDEX = INDEX Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
'+-----------------------------+
'+  iB1_DenSyu
'+-----------------------------+
Private Sub iB1_DenSyu_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_DenSyu_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    B1lb_DenSyu = ""
    If H1Op_Torkb(0).Value Then
        gStr = "ì`ï[éÌóﬁéÛíç"
    Else
        gStr = "ì`ï[éÌóﬁî≠íç"
    End If
    If MeisyoGet(gStr, iB1_DenSyu) Then
        B1lb_DenSyu = SYM_Meisyo
    End If
    
    If iB1_DenSyu = "3" Or iB1_DenSyu = "4" Then
        B1Chk_PrtkbS.Value = 1: Call B1Chk_PrtkbS_Click
        B1Chk_PrtkbJ.Value = 1: Call B1Chk_PrtkbJ_Click
    End If

End Sub
'''2007/12/11DEL'+-----------------------------+
'''2007/12/11DEL'+  B1Comb_NouhinÅ@î[ïièëå`ë‘
'''2007/12/11DEL'+-----------------------------+
'''2007/12/11DELPrivate Sub B1Comb_Nouhin_GotFocus()
'''2007/12/11DEL    B1Comb_Nouhin.BackColor = gIGotSel
'''2007/12/11DELEnd Sub
'''2007/12/11DELPrivate Sub B1Comb_Nouhin_LostFocus()
'''2007/12/11DEL    B1Comb_Nouhin.BackColor = gILostSel
'''2007/12/11DELEnd Sub
'+-----------------------------+
'+  B1Chk_PrtkbNÅ@î[ïièëî≠çsãÊï™
'+-----------------------------+
Private Sub B1Chk_PrtkbN_Click()
    If B1Chk_PrtkbN.Value = 0 Then
       B1Chk_PrtkbN.Caption = "î≠çsóL"
    Else
       B1Chk_PrtkbN.Caption = "î≠çsñ≥"
    End If
End Sub
'+-----------------------------+
'+  B1Chk_PrtkbSÅ@éxããèëî≠çsãÊï™
'+-----------------------------+
Private Sub B1Chk_PrtkbS_Click()
    If B1Chk_PrtkbS.Value = 0 Then
       B1Chk_PrtkbS.Caption = "î≠çsóL"
    Else
       B1Chk_PrtkbS.Caption = "î≠çsñ≥"
    End If
End Sub
'+-----------------------------+
'+  B1Chk_PrtkbJÅ@éÛóÃèëî≠çsãÊï™
'+-----------------------------+
Private Sub B1Chk_PrtkbJ_Click()
    If B1Chk_PrtkbJ.Value = 0 Then
       B1Chk_PrtkbJ.Caption = "î≠çsóL"
    Else
       B1Chk_PrtkbJ.Caption = "î≠çsñ≥"
    End If
End Sub
'''2007/12/11DEL'+-----------------------------+
'''2007/12/11DEL'+  B1Comb_JyuryoÅ@éÛóÃèëå`ë‘
'''2007/12/11DEL'+-----------------------------+
'''2007/12/11DELPrivate Sub B1Comb_Jyuryo_GotFocus()
'''2007/12/11DEL    B1Comb_Jyuryo.BackColor = gIGotSel
'''2007/12/11DELEnd Sub
'''2007/12/11DELPrivate Sub B1Comb_Jyuryo_LostFocus()
'''2007/12/11DEL    B1Comb_Jyuryo.BackColor = gILostSel
'''2007/12/11DELEnd Sub
'+-----------------------------+
'+  B1Comb_HNaijiÅ@î≠íçì‡é¶
'+-----------------------------+
Private Sub B1Comb_HNaiji_GotFocus()
    B1Comb_HNaiji.BackColor = gIGotSel
End Sub
Private Sub B1Comb_HNaiji_LostFocus()
    B1Comb_HNaiji.BackColor = gILostSel
End Sub
'+-----------------------------+
'+  B1Comb_MeaisaiÅ@ñæç◊èëî≠çs
'+-----------------------------+
Private Sub B1Comb_Meaisai_GotFocus()
    B1Comb_Meaisai.BackColor = gIGotSel
End Sub
Private Sub B1Comb_Meaisai_LostFocus()
    B1Comb_Meaisai.BackColor = gILostSel
End Sub
'+-----------------------------+
'+  B1Chk_JuryoÅ@éÛóÃèëî≠çs
'+-----------------------------+
Private Sub B1Chk_Juryo_Click()
    If H1Op_Torkb(0).Value Then
        If B1Chk_Juryo.Value = 0 Then
           B1Chk_Juryo.Caption = "éÛíçéû"
        Else
           B1Chk_Juryo.Caption = "èoâ◊éû"
        End If
    Else
        If B1Chk_Juryo.Value = 0 Then
           B1Chk_Juryo.Caption = "î≠íçéû"
        Else
           B1Chk_Juryo.Caption = "éÛóÃéû"
        End If
    End If
End Sub
'+-----------------------------+
'+  iB1_JYokin
'+-----------------------------+
Private Sub iB1_JYokin_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_JYokin_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    B1lb_JYokin = ""
    If MeisyoGet("óaã‡ãÊï™", iB1_JYokin) Then
        B1lb_JYokin = SYM_Meisyo
    End If
    
End Sub
'+-----------------------------+
'+  iB1_STHYmd
'+-----------------------------+
Private Sub iB1_STHYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_STHYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

End Sub
'+-----------------------------+
'+  FAXëóêMêÊ
'+-----------------------------+
Private Sub iB1_SendFax_GotFocus(INDEX As Integer)
    PB_Look.Tag = Format(INDEX, "00") & ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_SendFax_LostFocus(INDEX As Integer)
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    B1lb_SendFaxTor(INDEX) = ""
    B1lb_SendFax(INDEX) = ""

    If TorNmGet(iB1_SendFax(INDEX), 8) Then
        B1lb_SendFaxTor(INDEX) = TRM_RName
        B1lb_SendFax(INDEX) = TRM_Fax
    End If
End Sub
'+-----------------------------+
'+  MAILëóêMêÊ
'+-----------------------------+
Private Sub iB1_SendMail_GotFocus(INDEX As Integer)
    PB_Look.Tag = Format(INDEX, "00") & ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_SendMail_LostFocus(INDEX As Integer)
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    B1lb_SendMailTor(INDEX) = ""
    B1lb_SendMail(INDEX) = ""

    If TorNmGet(iB1_SendMail(INDEX), 8) Then
        B1lb_SendMailTor(INDEX) = TRM_RName
        B1lb_SendMail(INDEX) = TRM_EMail
    End If
End Sub
'============================================ 2009/07/15 Start
'+-----------------------------+
'+  B1Chk_Sijisyo  éwé¶èë
'+-----------------------------+
Private Sub B1Chk_Sijisyo_Click()
    If B1Chk_Sijisyo.Value = 1 Then
       B1Chk_Sijisyo.Caption = "î≠çsñ≥"
    Else
       B1Chk_Sijisyo.Caption = "î≠çsóL"
    End If
End Sub
'+-----------------------------+
'+  B1Chk_NipoÅ@ì˙ïÒ
'+-----------------------------+
Private Sub B1Chk_Nipo_Click()
    If B1Chk_Nipo.Value = 1 Then
       B1Chk_Nipo.Caption = "î≠çsñ≥"
    Else
       B1Chk_Nipo.Caption = "î≠çsóL"
    End If
End Sub
'+-----------------------------+
'+  B1Chk_FukaÅ@ïââ◊
'+-----------------------------+
Private Sub B1Chk_Fuka_Click()
    If B1Chk_Fuka.Value = 1 Then
       B1Chk_Fuka.Caption = "ÇµÇ»Ç¢"
    Else
       B1Chk_Fuka.Caption = "Ç∑ÇÈ"
    End If
End Sub
'+-----------------------------+
'+  B1Chk_PSChkÅ@ïââ◊
'+-----------------------------+
Private Sub B1Chk_PSChk_Click()
    If B1Chk_PSChk.Value = 1 Then
       B1Chk_PSChk.Caption = "ÉXÉ|ÉbÉg"
    Else
       B1Chk_PSChk.Caption = "ÉvÉåÉX"
    End If
End Sub
'============================================ 2009/07/15 End

'+---------------------------------------------------------+
'+          ÉOÉäÉbÉhï“èWÉGÉäÉA
'+---------------------------------------------------------+
'+---------------------------------------------+
'+  iB1_Ukeir    éÛì¸
'+---------------------------------------------+
Private Sub iB1_Ukeir_GotFocus()
    Set ActObj = iB1_Ukeir
End Sub
Private Sub iB1_Ukeir_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Ukeir_LostFocus()
    With vsGrid
        ActRow = CisFun.Val2(iB1_Ukeir.Tag)
        If ActRow > .Rows - 1 Then Exit Sub
        .TextMatrix(ActRow, 1) = RTrim(iB1_Ukeir)
            
        If Trim(iB1_Ukeir) = "" Then
            For gInt = 2 To .Cols - 1
                .TextMatrix(ActRow, gInt) = ""
            Next gInt
            iB1_UName = ""
            iB1_Cycle1 = ""
            iB1_Cycle2 = ""
            iB1_Cycle3 = ""
            ChkB1_Kara.Value = 0
            ChkB1_YoubiHC.Value = 0
            ChkB1_Youbi(0).Value = 0
            ChkB1_Youbi(1).Value = 0
            ChkB1_Youbi(2).Value = 0
            ChkB1_Youbi(3).Value = 0
            ChkB1_Youbi(4).Value = 0
        End If
        
    End With
End Sub
'+---------------------------------------------+
'+  iB1_UName    éÛì¸ñºèÃ
'+---------------------------------------------+
Private Sub iB1_UName_GotFocus()
    Set ActObj = iB1_UName
End Sub
Private Sub iB1_UName_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_UName_LostFocus()
    With vsGrid
        .TextMatrix(CisFun.Val2(iB1_Ukeir.Tag), 2) = iB1_UName
    End With
End Sub
'+---------------------------------------------+
'+  iB1_Cycle1    ÉTÉCÉNÉã 1
'+---------------------------------------------+
Private Sub iB1_Cycle1_GotFocus()
    Set ActObj = iB1_Cycle1
End Sub
Private Sub iB1_Cycle1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Cycle1_LostFocus()
    With vsGrid
        ActRow = CisFun.Val2(iB1_Ukeir.Tag)
        If Trim(iB1_Ukeir) = "" Then
           .TextMatrix(ActRow, 3) = ""
           Exit Sub
        End If
        If ActRow > .Rows - 1 Then Exit Sub
        .TextMatrix(ActRow, 11) = RTrim(iB1_Cycle1)
        .TextMatrix(ActRow, 3) = CycleFormat(ActRow)
    End With
End Sub
'+---------------------------------------------+
'+  iB1_Cycle2    ÉTÉCÉNÉã 2
'+---------------------------------------------+
Private Sub iB1_Cycle2_GotFocus()
    Set ActObj = iB1_Cycle2
End Sub
Private Sub iB1_Cycle2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Cycle2_LostFocus()
    With vsGrid
        ActRow = CisFun.Val2(iB1_Ukeir.Tag)
        If Trim(iB1_Ukeir) = "" Then
           .TextMatrix(ActRow, 3) = ""
           Exit Sub
        End If
        If ActRow > .Rows - 1 Then Exit Sub
        .TextMatrix(ActRow, 12) = RTrim(iB1_Cycle2)
        .TextMatrix(ActRow, 3) = CycleFormat(ActRow)
    End With
End Sub
'+---------------------------------------------+
'+  iB1_Cycle1    ÉTÉCÉNÉã 3
'+---------------------------------------------+
Private Sub iB1_Cycle3_GotFocus()
    Set ActObj = iB1_Cycle3
End Sub
Private Sub iB1_Cycle3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Cycle3_LostFocus()
    With vsGrid
        ActRow = CisFun.Val2(iB1_Ukeir.Tag)
        If Trim(iB1_Ukeir) = "" Then
           .TextMatrix(ActRow, 3) = ""
           Exit Sub
        End If
        If ActRow > .Rows - 1 Then Exit Sub
        .TextMatrix(ActRow, 13) = RTrim(iB1_Cycle1)
        .TextMatrix(ActRow, 3) = CycleFormat(ActRow)
    End With
End Sub
Private Function CycleFormat(editRow As Integer, Optional EditChr As String = "-") As String
    CycleFormat = ""
    With vsGrid
        If Trim(.TextMatrix(editRow, 11)) <> "" Or Trim(.TextMatrix(editRow, 12)) <> "" Or _
           Trim(.TextMatrix(editRow, 13)) <> "" Then
           CycleFormat = .TextMatrix(editRow, 11) & EditChr & _
                        Format(CisFun.Val2(.TextMatrix(editRow, 12)), "00") & EditChr & _
                        Format(CisFun.Val2(.TextMatrix(editRow, 13)), "00")
        End If
    End With
End Function
'+---------------------------------------------+
'+  ChkB1_Kara    ãÛñæç◊
'+---------------------------------------------+
Private Sub ChkB1_Kara_Click()
    If ChkB1_Kara.Value = 0 Then
        ChkB1_Kara.Caption = "óL"
    Else
        ChkB1_Kara.Caption = "ñ≥"
    End If
End Sub
Private Sub ChkB1_Kara_GotFocus()
    ChkB1_Kara.BackColor = gIGotSel
    Set ActObj = ChkB1_Kara
End Sub
Private Sub ChkB1_Kara_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub ChkB1_Kara_LostFocus()
    ChkB1_Kara.BackColor = gILostSel
    With vsGrid
        ActRow = CisFun.Val2(iB1_Ukeir.Tag)
        If Trim(iB1_Ukeir) = "" Then
           .TextMatrix(ActRow, 4) = ""
           Exit Sub
        End If
        If ActRow > .Rows - 1 Then Exit Sub
        If ActRow = 0 Then Exit Sub
        If ChkB1_Kara.Value = 0 Then
           .TextMatrix(ActRow, 4) = "óL"
        Else
           .TextMatrix(ActRow, 4) = "ñ≥"
        End If
    End With
End Sub
'+---------------------------------------------+
'+  ChkB1_YoubiHC    ójì˙î≠íç   2005.07.15
'+---------------------------------------------+
Private Sub ChkB1_YoubiHC_Click()
    If ChkB1_YoubiHC.Value = 0 Then
        ChkB1_YoubiHC.Caption = "ÇµÇ»Ç¢"
    Else
        ChkB1_YoubiHC.Caption = "Ç∑ÇÈ"
    End If
End Sub
Private Sub ChkB1_YoubiHC_GotFocus()
    ChkB1_YoubiHC.BackColor = gIGotSel
    Set ActObj = ChkB1_YoubiHC
End Sub
Private Sub ChkB1_YoubiHC_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub ChkB1_YoubiHC_LostFocus()
    ChkB1_YoubiHC.BackColor = gILostSel
    With vsGrid
        ActRow = CisFun.Val2(iB1_Ukeir.Tag)
        If Trim(iB1_Ukeir) = "" Then
           .TextMatrix(ActRow, 5) = ""
           Exit Sub
        End If
        If ActRow > .Rows - 1 Then Exit Sub
        If ActRow = 0 Then Exit Sub
        If ChkB1_YoubiHC.Value = 0 Then
           .TextMatrix(ActRow, 5) = "ÇµÇ»Ç¢"
        Else
           .TextMatrix(ActRow, 5) = "Ç∑ÇÈ"
        End If
    End With
End Sub
'+---------------------------------------------+
'+  ChkB1_Youbi    ójì˙ê›íË     2005.07.15
'+---------------------------------------------+
Private Sub ChkB1_Youbi_Click(INDEX As Integer)
    If ChkB1_Youbi(INDEX).Value = 0 Then
        ChkB1_Youbi(INDEX).Caption = ""
    Else
        ChkB1_Youbi(INDEX).Caption = "Åõ"
    End If
End Sub
Private Sub ChkB1_Youbi_GotFocus(INDEX As Integer)
    ChkB1_Youbi(INDEX).BackColor = gIGotSel
    Set ActObj = ChkB1_Youbi
    ChkB1_Youbi(0).Tag = Format(INDEX, "0")
End Sub
Private Sub ChkB1_Youbi_KeyDown(INDEX As Integer, KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub ChkB1_Youbi_LostFocus(INDEX As Integer)
    ChkB1_Youbi(INDEX).BackColor = gILostSel
    With vsGrid
        ActRow = CisFun.Val2(iB1_Ukeir.Tag)
        If Trim(iB1_Ukeir) = "" Then
           .TextMatrix(ActRow, 6 + INDEX) = ""
           Exit Sub
        End If
        If ActRow > .Rows - 1 Then Exit Sub
        If ActRow = 0 Then Exit Sub
        If ChkB1_Youbi(INDEX).Value = 0 Then
           .TextMatrix(ActRow, 6 + INDEX) = ""
        Else
           .TextMatrix(ActRow, 6 + INDEX) = "Åõ"
        End If
    End With
End Sub
'=========================================================== 2009/01/06 Start
'+---------------------------------------------------------+
'+          éÊà¯êÊCDïœä∑ì¸óÕÉOÉäÉbÉhï“èWÉGÉäÉA
'+---------------------------------------------------------+
'+---------------------------------------------+
'+  iB1_MotoCD    ïœä∑å≥éÊà¯êÊCD
'+---------------------------------------------+
Private Sub iB1_MotoCD_GotFocus()
    Set ActObj = iB1_MotoCD
End Sub
Private Sub iB1_MotoCD_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange1(KeyCode)
End Sub
Private Sub iB1_MotoCD_LostFocus()
    With vsGrid1
        ActRow = CisFun.Val2(iB1_MotoCD.Tag)
        If ActRow > .Rows - 1 Then Exit Sub
        .TextMatrix(ActRow, 1) = RTrim(iB1_MotoCD)
            
        If Trim(iB1_MotoCD) = "" Then
            For gInt = 2 To .Cols - 1
                .TextMatrix(ActRow, gInt) = ""
            Next gInt
            iB1_MotoUke = ""
            iB1_ChgUke = ""
        End If
        
    End With
End Sub
'+---------------------------------------------+
'+  iB1_MotoUke    ïœä∑å≥éÛì¸
'+---------------------------------------------+
Private Sub iB1_MotoUke_GotFocus()
    Set ActObj = iB1_MotoUke
End Sub
Private Sub iB1_MotoUke_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange1(KeyCode)
End Sub
Private Sub iB1_MotoUke_LostFocus()
    With vsGrid1
        .TextMatrix(CisFun.Val2(iB1_MotoCD.Tag), 2) = iB1_MotoUke
    End With
End Sub
'+---------------------------------------------+
'+  iB1_ChgUke    ïœä∑êÊéÛì¸
'+---------------------------------------------+
Private Sub iB1_ChgUke_GotFocus()
    Set ActObj = iB1_ChgUke
End Sub
Private Sub iB1_ChgUke_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange1(KeyCode)
End Sub
Private Sub iB1_ChgUke_LostFocus()
    With vsGrid1
        .TextMatrix(CisFun.Val2(iB1_MotoCD.Tag), 2) = iB1_ChgUke
    End With
End Sub
'+---------------------------------------------+
'+  iB1_HenKeiyaku    å_ñÒêÊ
'+---------------------------------------------+
Private Sub iB1_HenKeiyaku_GotFocus()
    Set ActObj = iB1_HenKeiyaku
End Sub
Private Sub iB1_HenKeiyaku_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange1(KeyCode)
End Sub
Private Sub iB1_HenKeiyaku_LostFocus()
    With vsGrid1
        .TextMatrix(CisFun.Val2(iB1_HenKeiyaku.Tag), 2) = iB1_HenKeiyaku
    End With
End Sub
'=========================================================== 2009/01/06 End
'=========================================================== 2009/07/15 Start
'+---------------------------------------------------------+
'+          çHíˆèÓïÒì¸óÕÉOÉäÉbÉhï“èWÉGÉäÉA
'+---------------------------------------------------------+
'+---------------------------------------------+
'+  iB1_Kikai    ã@äB
'+---------------------------------------------+
Private Sub iB1_Kikai_GotFocus()
    Set ActObj = iB1_Kikai
End Sub
Private Sub iB1_Kikai_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange2(KeyCode)
End Sub
Private Sub iB1_Kikai_LostFocus()
    If RTrim(iB1_Kikai) <> "" Then
       Call MeisyoGet("ã@äB", RTrim(iB1_Kikai))
       B1lb_KikaiNM = RTrim(SYM_Meisyo)
    End If
    
    With VsGrid2
        ActRow = CisFun.Val2(iB1_Kikai.Tag)
        If ActRow > .Rows - 1 Then Exit Sub
        .TextMatrix(ActRow, 1) = RTrim(iB1_Kikai)
        .TextMatrix(ActRow, 2) = RTrim(B1lb_KikaiNM)
            
        If Trim(iB1_Kikai) = "" And Trim(iB1_KikaiNo) = "" Then
            For gInt = 2 To .Cols - 1
                .TextMatrix(ActRow, gInt) = ""
            Next gInt
            B1lb_KikaiNM = ""
            iB1_KikaiNo = ""
            B1lb_KikaiNoNM = ""
            iB1_KadoTime = ""
            iB1_Daisu = ""
            iB1_DaiKikai = ""
            B1lb_DaiKikaiNM = ""
            iB1_DaiKikaiNo = ""
            B1lb_DaiKikaiNoNM = ""
        End If
        
    End With

End Sub
'+---------------------------------------------+
'+  iB1_KikaiNo    ã@äBáÇ
'+---------------------------------------------+
Private Sub iB1_KikaiNo_GotFocus()
    Set ActObj = iB1_KikaiNo
End Sub
Private Sub iB1_KikaiNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange2(KeyCode)
End Sub
Private Sub iB1_KikaiNo_LostFocus()
    If RTrim(iB1_KikaiNo) <> "" Then
       Call MeisyoGet("ã@äBNO", RTrim(iB1_KikaiNo))
       B1lb_KikaiNoNM = RTrim(SYM_Meisyo)
    End If

    With VsGrid2
        ActRow = CisFun.Val2(iB1_Kikai.Tag)
        If ActRow > .Rows - 1 Then Exit Sub
        .TextMatrix(ActRow, 3) = RTrim(iB1_KikaiNo)
        .TextMatrix(ActRow, 4) = RTrim(B1lb_KikaiNoNM)

        If Trim(iB1_Kikai) = "" And Trim(iB1_KikaiNo) = "" Then
            For gInt = 2 To .Cols - 1
                .TextMatrix(ActRow, gInt) = ""
            Next gInt
            B1lb_KikaiNM = ""
            iB1_KikaiNo = ""
            B1lb_KikaiNoNM = ""
            iB1_KadoTime = ""
            iB1_Daisu = ""
            iB1_DaiKikai = ""
            B1lb_DaiKikaiNM = ""
            iB1_DaiKikaiNo = ""
            B1lb_DaiKikaiNoNM = ""
        End If

    End With

'    With VsGrid2
'        .TextMatrix(CisFun.Val2(iB1_Kikai.Tag), 2) = iB1_KadoTime
'    End With
End Sub
'+---------------------------------------------+
'+  iB1_KadoTime    â“ì≠éûä‘
'+---------------------------------------------+
Private Sub iB1_KadoTime_GotFocus()
    Set ActObj = iB1_KadoTime
End Sub
Private Sub iB1_KadoTime_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange2(KeyCode)
End Sub
Private Sub iB1_KadoTime_LostFocus()
    With VsGrid2
        .TextMatrix(CisFun.Val2(iB1_Kikai.Tag), 5) = iB1_KadoTime
    End With
End Sub
'+---------------------------------------------+
'+  iB1_Daisu    ë‰êî
'+---------------------------------------------+
Private Sub iB1_Daisu_GotFocus()
    Set ActObj = iB1_Daisu
End Sub
Private Sub iB1_Daisu_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange2(KeyCode)
End Sub
Private Sub iB1_Daisu_LostFocus()
    With VsGrid2
        .TextMatrix(CisFun.Val2(iB1_Kikai.Tag), 6) = iB1_Daisu
    End With
End Sub
'+---------------------------------------------+
'+  iB1_DaiKikai    ë„ë÷ã@äB
'+---------------------------------------------+
Private Sub iB1_DaiKikai_GotFocus()
    Set ActObj = iB1_DaiKikai
End Sub
Private Sub iB1_DaiKikai_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange2(KeyCode)
End Sub
Private Sub iB1_DaiKikai_LostFocus()
    If RTrim(iB1_DaiKikai) <> "" Then
       Call MeisyoGet("ã@äB", RTrim(iB1_DaiKikai))
       B1lb_DaiKikaiNM = RTrim(SYM_Meisyo)
    End If
    
    With VsGrid2
        .TextMatrix(CisFun.Val2(iB1_Kikai.Tag), 7) = iB1_DaiKikai
        .TextMatrix(CisFun.Val2(iB1_Kikai.Tag), 8) = B1lb_DaiKikaiNM
    End With
End Sub
'+---------------------------------------------+
'+  iB1_DaiKikaiNo    ë„ë÷ã@äBáÇ
'+---------------------------------------------+
Private Sub iB1_DaiKikaiNo_GotFocus()
    Set ActObj = iB1_DaiKikaiNo
End Sub
Private Sub iB1_DaiKikaiNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange2(KeyCode)
End Sub
Private Sub iB1_DaiKikaiNo_LostFocus()
    If RTrim(iB1_DaiKikaiNo) <> "" Then
       Call MeisyoGet("ã@äBNO", RTrim(iB1_DaiKikaiNo))
       B1lb_DaiKikaiNoNM = RTrim(SYM_Meisyo)
    End If
    
    With VsGrid2
        .TextMatrix(CisFun.Val2(iB1_Kikai.Tag), 9) = iB1_DaiKikaiNo
        .TextMatrix(CisFun.Val2(iB1_Kikai.Tag), 10) = B1lb_DaiKikaiNoNM
    End With
End Sub
'=========================================================== 2009/07/15 End
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
    PB_Look.Visible = H1Mode
    PB_ADD.Visible = H1Mode
    PB_MNT.Visible = H1Mode
    PB_DEL.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    
    PB_Kotei.Visible = B1Mode                   '2009/07/15
    
    '================================================= 2009/01/06 Start
    If ProcHB = "B1" Then
       If H1Op_Torkb(0) Then
          PB_CHG.Visible = True
       Else
          PB_CHG.Visible = False
       End If
    Else
       PB_CHG.Visible = False
    End If
    '================================================= 2009/01/06 End
    
    '=========================================== *** 2009/07/06 Start
    If Pro = "B1" And H1Op_Torkb(0).Value = True Then
       If B1Chk_Torkb(0).Value = 1 Then
          Back_Maekotei(0).Visible = False
          Back_Maekotei(1).Visible = False
          Back_NJigyo.Visible = True
          Back_SJigyo.Visible = True
          Back_Syukaba.Visible = True
       Else
          Back_Maekotei(0).Visible = False
          Back_Maekotei(1).Visible = True
          Back_NJigyo.Visible = True
          Back_SJigyo.Visible = False
          Back_Syukaba.Visible = True
       End If
    Else
       Back_Maekotei(0).Visible = True
       Back_Maekotei(1).Visible = False
       Back_NJigyo.Visible = False
       Back_SJigyo.Visible = False
       Back_Syukaba.Visible = False
    End If
    '=========================================== *** 2009/07/06 End
    
    Call GridTextEnabled(B1Mode)
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
'*--------------------------------------------------<< 2004/06/30 Insert Start >>
'           If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
'*--------------------------------------------------<< 2004/06/30 Insert End >>
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
                .BackColor = mLostColor
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
'            If HeadBody = "B1" Then
'               If .Name Like "Back_*" Then .Visible = True
'            End If
        End With
    Next gObj
    If HeadBody = "B" Or HeadBody = "B1" Then
        CisVsGrid.Clear
        
        CisVsGrid1.Clear                    '2009/07/15
        CisVsGrid2.Clear                    '2009/07/15
    End If
    If HeadBody = "BX" Then
        B1Op_Keisan(mIKeisan).Value = True
        B1Op_Marume(mIMarume).Value = True
        B1Op_Hei(mIHei).Value = True
        '------------------------------------------------------ ***** 2008/06/11 Start
        B1Op_Kurai1(mIKurai).Value = True
        B1Op_Kurai2(mIKurai).Value = True
        '------------------------------------------------------ ***** 2008/06/11 End
'''2007/12/11DEL        CisComb1.ComboSet
        CisComb2.ComboSet
        CisComb4.ComboSet
        B1Chk_Juryo.Value = 1
        If H1Op_Torkb(0).Value Then
'''2007/12/11DEL            CisComb3.ComboName = "î[ïièëå`ë‘éÛíç"
           B1Chk_Juryo.Caption = "èoâ◊éû"
        Else
'''2007/12/11DEL            CisComb3.ComboName = "î[ïièëå`ë‘î≠íç"
           B1Chk_Juryo.Caption = "éÛóÃéû"
        End If
'''2007/12/11DEL        CisComb3.ComboSet
    
        '========================================= 2009/07/15 Start
        CisComb5.ComboSet
        CisComb6.ComboSet
        '========================================= 2009/07/15 End
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
            Call VsGridToText(1)
            ProcHB = "B1"
            Call DispChange(ProcHB)
            
            If B1Chk_Torkb(1).Value = 0 Then
               PB_Kotei.Visible = False
            End If
            
            For Each gObj In B1Chk_Torkb
                If gObj.Enabled Then
                   gObj.SetFocus
                   
                   Exit For
                End If
            Next gObj
        End If
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If B1Chk_Torkb(1).Value = 1 And ProcHB = "B1" And B1_Area1.Enabled = True Then
           PB_Kotei.Visible = False
        End If
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then
            For Each gObj In B1Chk_Torkb
                If gObj.Enabled Then
                   gObj.SetFocus
                   Exit For
                End If
            Next gObj
            GoTo ReturnPress_Ed
        End If
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
    
    If Trim(iH1_Torcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    éÊà¯êÊÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    End If

' É}ÉXÉ^ì¸óÕ SELECTï∂ æØƒ
    gSL_Select = " SELECT * "
    gSL_Select = gSL_Select & " FROM éÊà¯êÊÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " WHERE éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & RTrim(iH1_Kojyo) & "'"
    gSL_Select = gSL_Select & "   And éÊà¯êÊãÊï™ = " & CisFun.Val2(H1Op_Torkb(0).Tag)
    If DBInput("INV") Then
       If SyoriKB = "ADD" Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    éÊà¯êÊÉ}ÉXÉ^ìoò^çœ            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          iH1_Torcd.SetFocus
          PB_Kotei.Visible = False                  '2009/07/15
          Exit Function
       End If
    Else
       If SyoriKB <> "ADD" Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    éÊà¯êÊÉ}ÉXÉ^ñ¢ìoò^            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          iH1_Torcd.SetFocus
          PB_Kotei.Visible = False                  '2009/07/15
          Exit Function
       End If
    End If
    
'    If CisFun.Val2(H1Op_Torkb(0).Tag) = 0 Then
'       B1Chk_Torkb(1).Value = 0
'       B1Chk_Torkb(1).Enabled = False
'    Else
'       B1Chk_Torkb(1).Enabled = True
'    End If
    Call BodyEnbled
    
    '=========================================== 2009/07/15 Start
    If SyoriKB <> "DEL" Then
       If SyoriKB = "ADD" Then PB_Kotei.Visible = False
       If SyoriKB = "MNT" Then
          If B1Chk_Torkb(1).Value = 1 Then
             PB_Kotei.Visible = True
          Else
             PB_Kotei.Visible = False
          End If
       End If
    End If
    '=========================================== 2009/07/15 End
    
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
    Head1Chk = True
End Function
'****************************
'*      Body1¡™Ø∏èàóù
'****************************
Private Function Body1Chk() As Boolean
    
    Call TextToVsGrid(Val(iB1_Ukeir.Tag))
    
    Body1Chk = False
    
    If iB1_Tornm = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    ê≥éÆñºèÃÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_Tornm.SetFocus
       Exit Function
    End If

    '============================================ 2008/12/26 Start
    If H1Op_Torkb(0) Then
       'ïœä∑å≥éÊà¯êÊCD,ë∂ç›¡™Ø∏
       If RTrim(iB1_ConvTorcd) & RTrim(iB1_ConvTorcdKj) <> "" Then
          gSL_Select = "SELECT ïœä∑å≥éÊà¯êÊCD FROM éÊà¯êÊÉ}ÉXÉ^ "
          gSL_Select = gSL_Select & " WHERE ïœä∑å≥éÊà¯êÊCD = '" & RTrim(iB1_ConvTorcd) & RTrim(iB1_ConvTorcdKj) & "'"
          gSL_Select = gSL_Select & " AND   éÊà¯êÊãÊï™ = 0 "
          If TRMRead(gSL_Select, 1) Then
             With CisFun
                  .MB_Lines = 5
                  .MB_MSG(2) = "    ì¸óÕÇ≥ÇÍÇΩ[ïœä∑å≥éÊà¯êÊ∫∞ƒﬁ]ÇÕä˘Ç…        "
                  .MB_MSG(3) = "    ë∂ç›ÇµÇƒÇ¢Ç‹Ç∑ÅB        "
                  .MB_MSG(4) = "  Å@Ç±ÇÃÇ‹Ç‹ÇÃílÇ≈ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH  "
                  .MB_Title = ""
                  .MB_Button = Yes_No
                  If Not .MBOX Then
                     iB1_ConvTorcd.SetFocus
                     Exit Function
                  End If
             End With
          End If
       End If
    End If
    '============================================ 2008/12/26 End

'    If iB1_KojyoNm.Visible And Trim(iB1_KojyoNm) = "" Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    çHèÍñºèÃÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
'            .MB_Title = ""
'            .MB_Button = Error
'            .MBOX
'        End With
'       iB1_KojyoNm.SetFocus
'       Exit Function
'    End If
    ' å_ñÒêÊ
    B1lb_KTorNm = ""
    If iB1_KTorcd = "" And iB1_KTorcd.Visible Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    å_ñÒêÊÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_KTorcd.SetFocus
       Exit Function
    End If
    If iB1_KTorcd.Visible Then
        If CisFun.Val2(H1Op_Torkb(0).Tag) = 0 Then
            If TorNmGet(iB1_KTorcd, 2) Then
                B1lb_KTorNm = TRM_RName
            Else
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    å_ñÒêÊñ¢ìoò^            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
               iB1_KTorcd.SetFocus
               Exit Function
            End If
        Else
            If TorNmGet(iB1_KTorcd, 5) Then
                B1lb_KTorNm = TRM_RName
            Else
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    å_ñÒêÊñ¢ìoò^            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
               iB1_KTorcd.SetFocus
               Exit Function
            End If
        End If
    End If
    
    '=================================================== *** 2009/07/06 Start
    If RTrim(iB1_NJigyoCD) = "" And RTrim(iB1_NJigyoKJ) <> "" Then
       With CisFun
           .MB_Lines = 4
           .MB_MSG(2) = "     î[ì¸êÊéñã∆ëÃÇì¸óÕÇµÇƒâ∫Ç≥Ç¢          "
           .MB_Title = ""
           .MB_Button = Error
           .MBOX
       End With
      iB1_NJigyoCD.SetFocus
      Exit Function
    End If
    If RTrim(iB1_NJigyoCD) <> "" And RTrim(iB1_NJigyoKJ) = "" Then
       With CisFun
           .MB_Lines = 4
           .MB_MSG(2) = "     î[ì¸êÊéñã∆ëÃçHãÊÇì¸óÕÇµÇƒâ∫Ç≥Ç¢          "
           .MB_Title = ""
           .MB_Button = Error
           .MBOX
       End With
      iB1_NJigyoKJ.SetFocus
      Exit Function
    End If
    '=================================================== *** 2009/07/06 End
    
    ' ì`ï[éÌóﬁ
    B1lb_DenSyu = "": SaveKeitai = ""
    If Trim(iB1_DenSyu) <> "" And iB1_DenSyu.Visible Then
        If H1Op_Torkb(0).Value Then
            gStr = "ì`ï[éÌóﬁéÛíç"
        Else
            gStr = "ì`ï[éÌóﬁî≠íç"
        End If
        If MeisyoGet(gStr, iB1_DenSyu) Then
            B1lb_DenSyu = SYM_Meisyo
            SaveKeitai = CisFun.Left2(SYM.âpêîéö1, 1)
        Else
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ì`ï[éÌóﬁñ¢ìoò^            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iB1_DenSyu.SetFocus
           Exit Function
        End If
        If H1Op_Torkb(1).Value Then
'''2007/12/11DEL            CisComb3.CodeGet
'''2007/12/11DEL            CisComb1.CodeGet
'''2007/12/11DEL            If RTrim(iB1_DenSyu) = "2" And (CisComb1.Code <> "2" Or CisComb3.Code <> "2") Then
'''2007/12/11DEL                With CisFun
'''2007/12/11DEL                    .MB_Lines = 4
'''2007/12/11DEL                    .MB_MSG(1) = "    ì`ï[éÌóﬁÇ™ éxããÇiÇ`ÇlÇ`ÇÃéûÅAî[ïièëå`ë‘ÅEéÛóÃèëå`ë‘ÇÕÉtÉ@ÉCÉãèoóÕå¿íËÇ≈Ç∑ÅB            "
'''2007/12/11DEL                    .MB_MSG(3) = "    î[ïièëå`ë‘ÅEéÛóÃèëå`ë‘ÇÉtÉ@ÉCÉãèoóÕÇ…ê›íËÇµÇ‹Ç∑ÅB            "
'''2007/12/11DEL                    .MB_Title = ""
'''2007/12/11DEL                    .MB_Button = Error
'''2007/12/11DEL                    .MBOX
'''2007/12/11DEL                End With
'''2007/12/11DEL                CisComb1.Code = "2"
'''2007/12/11DEL                CisComb1.TextGet
'''2007/12/11DEL                CisComb3.Code = "2"
'''2007/12/11DEL                CisComb3.TextGet
'''2007/12/11DEL            End If
'''2007/12/11DEL    '----<< É`ÉPÉbÉgí«â¡  >>-------------------------------------------------------------- 2005.01.31 (START)-------'
'''2007/12/11DEL            If RTrim(iB1_DenSyu) = "3" And CisComb1.Code <> "" Then
'''2007/12/11DEL                With CisFun
'''2007/12/11DEL                    .MB_Lines = 4
'''2007/12/11DEL                    .MB_MSG(1) = "    ì`ï[éÌóﬁÇ™ é©ããÉ`ÉPÉbÉgÇÃéûÅAéÛóÃèëå`ë‘ÇÕèoóÕñ≥Çµå¿íËÇ≈Ç∑ÅB            "
'''2007/12/11DEL                    .MB_MSG(3) = "    éÛóÃèëå`ë‘ÇèoóÕñ≥ÇµÇ…ê›íËÇµÇ‹Ç∑ÅB            "
'''2007/12/11DEL                    .MB_Title = ""
'''2007/12/11DEL                    .MB_Button = Error
'''2007/12/11DEL                    .MBOX
'''2007/12/11DEL                End With
'''2007/12/11DEL                CisComb1.Code = "0"
'''2007/12/11DEL                CisComb1.TextGet
'''2007/12/11DEL            End If
'''2007/12/11DEL            If RTrim(iB1_DenSyu) = "4" And (CisComb1.Code <> "" Or CisComb3.Code <> "2") Then
'''2007/12/11DEL                With CisFun
'''2007/12/11DEL                    .MB_Lines = 4
'''2007/12/11DEL                    .MB_MSG(1) = "    ì`ï[éÌóﬁÇ™ éxããÉ`ÉPÉbÉgÇÃéûÅAî[ïièëå`ë‘ÇÕÉtÉ@ÉCÉãèoóÕÅEéÛóÃèëå`ë‘ÇÕèoóÕñ≥Çµå¿íËÇ≈Ç∑ÅB            "
'''2007/12/11DEL                    .MB_MSG(3) = "    î[ïièëå`ë‘ÇÉtÉ@ÉCÉãèoóÕÅEéÛóÃèëå`ë‘ÇèoóÕñ≥ÇµÇ…ê›íËÇµÇ‹Ç∑ÅB            "
'''2007/12/11DEL                    .MB_Title = ""
'''2007/12/11DEL                    .MB_Button = Error
'''2007/12/11DEL                    .MBOX
'''2007/12/11DEL                End With
'''2007/12/11DEL                CisComb1.Code = "0"
'''2007/12/11DEL                CisComb1.TextGet
'''2007/12/11DEL                CisComb3.Code = "2"
'''2007/12/11DEL                CisComb3.TextGet
'''2007/12/11DEL            End If
'''2007/12/11DEL    '----<< É`ÉPÉbÉgí«â¡  >>-------------------------------------------------------------- 2005.01.31 (END  )-------'
            If (RTrim(iB1_DenSyu) = "3" Or RTrim(iB1_DenSyu) = "4") And (B1Chk_PrtkbS.Value = 0 Or B1Chk_PrtkbJ.Value = 0) Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(1) = "    ì`ï[éÌóﬁÇ™ É`ÉPÉbÉgÇÃéûÅAéxããèëî≠çsÅEéÛóÃèëî≠çsÇÕî≠çsñ≥Çµå¿íËÇ≈Ç∑ÅB            "
                    .MB_MSG(3) = "    éxããèëî≠çsÅEéÛóÃèëî≠çsÇî≠çsñ≥ÇµÇ…ê›íËÇµÇ‹Ç∑ÅB            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                B1Chk_PrtkbS.Value = 1
                B1Chk_PrtkbJ.Value = 1
            End If
        End If
'''2007/12/11DEL        If B1Comb_Nouhin.ListIndex = 0 And B1Comb_Jyuryo.ListIndex = 0 Then
'''2007/12/11DEL            With CisFun
'''2007/12/11DEL                .MB_Lines = 4
'''2007/12/11DEL                .MB_MSG(2) = "    î[ïièëå`ë‘ÅEéÛóÃèëå`ë‘ÇÇê›íËÇµÇƒâ∫Ç≥Ç¢ÅB            "
'''2007/12/11DEL                .MB_Title = ""
'''2007/12/11DEL                .MB_Button = Error
'''2007/12/11DEL                .MBOX
'''2007/12/11DEL            End With
'''2007/12/11DEL           B1Comb_Nouhin.SetFocus
'''2007/12/11DEL           Exit Function
'''2007/12/11DEL        End If
    
    End If
    If Trim(iB1_DenSyu) = "" Then
'''2007/12/11DEL        B1Comb_Nouhin.ListIndex = 0
'''2007/12/11DEL        B1Comb_Jyuryo.ListIndex = 0
        B1Chk_PrtkbN.Value = 1: B1Chk_PrtkbS.Value = 1: B1Chk_PrtkbJ.Value = 1
        B1Chk_Juryo.Value = 0
    End If
    ' óaã‡ãÊï™
    B1lb_JYokin = ""
    If Trim(iB1_JYokin) <> "" And Back_Ginko.Visible Then
        If MeisyoGet("óaã‡ãÊï™", iB1_JYokin) Then
            B1lb_JYokin = SYM_Meisyo
        Else
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    óaã‡ãÊï™ñ¢ìoò^            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iB1_JYokin.SetFocus
           Exit Function
        End If
    End If
    '+----------------------+
    '+   ëóêMèÓïÒ
    '+----------------------+
    For gCnt1 = 0 To 2
        If Trim(iB1_SendFax(gCnt1)) <> "" Then
            If Not TorNmGet(iB1_SendFax(gCnt1), 8) Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    ÉfÅ[É^ëóêMêÊèÓïÒÅiÇeÇ`ÇwÅjñ¢ìoò^            "
                    .MB_MSG(4) = "              " & StrConv(gCnt1 + 1, vbWide) & "çsñ⁄"
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
               iB1_SendFax(gCnt1).SetFocus
               Exit Function
            End If
            If Trim(TRM_Fax) = "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    ÉfÅ[É^ëóêMêÊèÓïÒÅiÇeÇ`ÇwÅjÇeÇ`Çwî‘çÜñ¢ìoò^            "
                    .MB_MSG(4) = "              " & StrConv(gCnt1 + 1, vbWide) & "çsñ⁄"
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
               iB1_SendFax(gCnt1).SetFocus
               Exit Function
            End If
         End If
    Next gCnt1
    For gCnt1 = 0 To 2
        If Trim(iB1_SendMail(gCnt1)) <> "" Then
            If Not TorNmGet(iB1_SendMail(gCnt1), 8) Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    ÉfÅ[É^ëóêMêÊèÓïÒÅie-MailÅjñ¢ìoò^            "
                    .MB_MSG(4) = "              " & StrConv(gCnt1 + 1, vbWide) & "çsñ⁄"
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
               iB1_SendMail(gCnt1).SetFocus
               Exit Function
            End If
            If Trim(TRM_EMail) = "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    ÉfÅ[É^ëóêMêÊèÓïÒÅie-MailÅje-Mailñ¢ìoò^            "
                    .MB_MSG(4) = "              " & StrConv(gCnt1 + 1, vbWide) & "çsñ⁄"
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
               iB1_SendMail(gCnt1).SetFocus
               Exit Function
            End If
         End If
    Next gCnt1
    
    Call CisComb4.CodeGet
    Call CisComb2.CodeGet
    
    If CisComb4.Code = "3" Or CisComb2.Code = "3" Then
       If Trim(iB1_JTFax) = "" And _
          Trim(iB1_SendFax(0)) = "" And _
          Trim(iB1_SendFax(1)) = "" And _
          Trim(iB1_SendFax(2)) = "" Then
          With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    ÇeÇ`Çw ñîÇÕ ÉfÅ[É^ëóêMêÊí«â¡èÓïÒÅiÇeÇ`ÇwÅjÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iB1_JTFax.SetFocus
          Exit Function
       End If
    End If
    If CisComb4.Code = "4" Or CisComb2.Code = "4" Then
       If Trim(iB1_JTEMail) = "" And _
          Trim(iB1_SendMail(0)) = "" And _
          Trim(iB1_SendMail(1)) = "" And _
          Trim(iB1_SendMail(2)) = "" Then
          With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    e-Mail ñîÇÕ ÉfÅ[É^ëóêMêÊí«â¡èÓïÒÅie-MailÅjÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iB1_JTEMail.SetFocus
          Exit Function
       End If
    End If
    
    '+--------------------+
    '+   éÛì¸èÓïÒ
    '+--------------------+
    With vsGrid
        For gCnt1 = 1 To .Rows - 1
            If Trim(.TextMatrix(gCnt1, 1)) <> "" Then
                'éÛì¸ÇÃèdï°É`ÉFÉbÉN
                For gCnt2 = gCnt1 + 1 To .Rows - 1
                    If Trim(.TextMatrix(gCnt2, 1)) <> "" Then
                        If Trim(.TextMatrix(gCnt1, 1)) = Trim(.TextMatrix(gCnt2, 1)) Then
                            Set gObj = iB1_Ukeir
                            gLong = gCnt2
                            With CisFun
                                .MB_Lines = 4
                                .MB_MSG(2) = "    " & StrConv(gCnt2, vbWide) & "çsñ⁄ éÛì¸Ç™èdï°ÇµÇƒÇ‹Ç∑ÅB            "
                                .MB_Title = ""
                                .MB_Button = Error
                                .MBOX
                            End With
                            GoTo Ukeir_Err
                        End If
                    End If
                Next gCnt2
                ' ÉTÉCÉNÉã
                If Trim(.TextMatrix(gCnt1, 11)) <> "" Or Trim(.TextMatrix(gCnt1, 12)) <> "" Or _
                   Trim(.TextMatrix(gCnt1, 13)) <> "" Then
                   If CisFun.Val2(.TextMatrix(gCnt1, 11)) = 0 Or _
                      CisFun.Val2(.TextMatrix(gCnt1, 12)) = 0 Or _
                      CisFun.Val2(.TextMatrix(gCnt1, 13)) = 0 Then
                      gLong = gCnt1
                      Set gObj = iB1_Cycle1
                      With CisFun
                           .MB_Lines = 4
                           .MB_MSG(2) = "  " & StrConv(gCnt1, vbWide) & "çsñ⁄  ÉTÉCÉNÉãÇê≥ÇµÇ≠ì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
                           .MB_Title = ""
                           .MB_Button = Error
                           .MBOX
                       End With
                       GoTo Ukeir_Err
                   End If
                End If
                '--------------------------------------------------------(2005.07.15) Start --
                ' ójì˙î≠íç
                If Trim(.TextMatrix(gCnt1, 5)) = "Ç∑ÇÈ" Then
                    gCnt3 = 0
                    For gCnt2 = 0 To 4
                        If Trim(.TextMatrix(gCnt1, 6 + gCnt2)) <> "" Then
                            gCnt3 = gCnt3 + 1
                        End If
                    Next gCnt2
                    If gCnt3 = 0 Then
                        Set gObj = ChkB1_Youbi(0)
                        gLong = gCnt1
                        With CisFun
                            .MB_Lines = 4
                            .MB_MSG(2) = "    " & StrConv(gCnt1, vbWide) & "çsñ⁄ ójì˙ÇÃê›íËÇ™Ç†ÇËÇ‹ÇπÇÒÅB            "
                            .MB_Title = ""
                            .MB_Button = Error
                            .MBOX
                        End With
                        gLong = gCnt1
                        Set gObj = ChkB1_Youbi(0)
                        GoTo Ukeir_Err
                    End If
                    If CisFun.Val2(.TextMatrix(gCnt1, 12)) <> 1 Then
                      With CisFun
                           .MB_Lines = 4
                           .MB_MSG(2) = "  " & StrConv(gCnt1, vbWide) & "çsñ⁄  ójì˙î≠íçÇÃéûÅAÉTÉCÉNÉãÅiâÒÅjÇÕÇPÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
                           .MB_Title = ""
                           .MB_Button = Error
                           .MBOX
                       End With
                       gLong = gCnt1
                       Set gObj = iB1_Cycle2
                       GoTo Ukeir_Err
                    End If
                    
                End If
                '--------------------------------------------------------(2005.07.15) End --
            End If
        Next gCnt1
    End With
    
    
    
    '=============================================== *** 2009/07/15 Start
    '+--------------------+
    '+   çHíˆèÓïÒ
    '+--------------------+
    If H1Op_Torkb(1) Then
        With VsGrid2
            For gCnt1 = 1 To .Rows - 1
                If Trim(.TextMatrix(gCnt1, 1)) <> "" Or Trim(.TextMatrix(gCnt1, 3)) <> "" Then
                    'ã@äB ¡™Ø∏
                    If Trim(.TextMatrix(gCnt1, 1)) <> "" Then
                        If Not MeisyoGet("ã@äB", RTrim(.TextMatrix(gCnt1, 1))) Then
                            Back_Kotei.Visible = True
                            B1Comb_Bumon1.SetFocus
                            PNL_PF.Visible = False
                            B1_Area1.Enabled = False
                            
                            Set gObj = iB1_Kikai
                            gLong = gCnt1
                            With CisFun
                                .MB_Lines = 4
                                .MB_MSG(2) = "    " & StrConv(gCnt1, vbWide) & "çsñ⁄ [ã@äB]Ç™ñ¢ìoò^Ç≈Ç∑ÅB            "
                                .MB_Title = ""
                                .MB_Button = Error
                                .MBOX
                            End With
                        
                            GoTo Ukeir_Err
                        End If
                    End If
                    'ã@äBNO ¡™Ø∏
                    If Trim(.TextMatrix(gCnt1, 3)) <> "" Then
                        If Not MeisyoGet("ã@äBNO", RTrim(.TextMatrix(gCnt1, 3))) Then
                            Back_Kotei.Visible = True
                            B1Comb_Bumon1.SetFocus
                            PNL_PF.Visible = False
                            B1_Area1.Enabled = False
                            
                            Set gObj = iB1_KikaiNo
                            gLong = gCnt1
                            With CisFun
                                .MB_Lines = 4
                                .MB_MSG(2) = "    " & StrConv(gCnt1, vbWide) & "çsñ⁄ [ã@äBáÇ]Ç™ñ¢ìoò^Ç≈Ç∑ÅB            "
                                .MB_Title = ""
                                .MB_Button = Error
                                .MBOX
                            End With
                        
                            GoTo Ukeir_Err
                        End If
                    End If
                    'ë„ë÷ã@äB ¡™Ø∏
                    If Trim(.TextMatrix(gCnt1, 7)) <> "" Then
                        If Not MeisyoGet("ã@äB", RTrim(.TextMatrix(gCnt1, 7))) Then
                            Back_Kotei.Visible = True
                            B1Comb_Bumon1.SetFocus
                            PNL_PF.Visible = False
                            B1_Area1.Enabled = False
                            
                            Set gObj = iB1_DaiKikai
                            gLong = gCnt1
                            With CisFun
                                .MB_Lines = 4
                                .MB_MSG(2) = "    " & StrConv(gCnt1, vbWide) & "çsñ⁄ [ë„ë÷ã@äB]Ç™ñ¢ìoò^Ç≈Ç∑ÅB            "
                                .MB_Title = ""
                                .MB_Button = Error
                                .MBOX
                            End With
                        
                            GoTo Ukeir_Err
                        End If
                    End If
                    'ë„ë÷ã@äBNO ¡™Ø∏
                    If Trim(.TextMatrix(gCnt1, 9)) <> "" Then
                        If Not MeisyoGet("ã@äBNO", RTrim(.TextMatrix(gCnt1, 9))) Then
                            Back_Kotei.Visible = True
                            B1Comb_Bumon1.SetFocus
                            PNL_PF.Visible = False
                            B1_Area1.Enabled = False
                            
                            Set gObj = iB1_DaiKikaiNo
                            gLong = gCnt1
                            With CisFun
                                .MB_Lines = 4
                                .MB_MSG(2) = "    " & StrConv(gCnt1, vbWide) & "çsñ⁄ [ë„ë÷ã@äBáÇ]Ç™ñ¢ìoò^Ç≈Ç∑ÅB            "
                                .MB_Title = ""
                                .MB_Button = Error
                                .MBOX
                            End With
                        
                            GoTo Ukeir_Err
                        End If
                    End If
                    
                    'ã@äB/ã@äBáÇÇÃèdï°É`ÉFÉbÉN
                    For gCnt2 = gCnt1 + 1 To .Rows - 1
                        If Trim(.TextMatrix(gCnt2, 1)) <> "" Or Trim(.TextMatrix(gCnt2, 3)) <> "" Then
                            If Trim(.TextMatrix(gCnt1, 1)) & Trim(.TextMatrix(gCnt1, 3)) = Trim(.TextMatrix(gCnt2, 1)) & Trim(.TextMatrix(gCnt2, 3)) Then
                                Back_Kotei.Visible = True
                                B1Comb_Bumon1.SetFocus
                                PNL_PF.Visible = False
                                B1_Area1.Enabled = False
                                
                                Set gObj = iB1_Kikai
                                gLong = gCnt2
                                With CisFun
                                    .MB_Lines = 4
                                    .MB_MSG(2) = "    " & StrConv(gCnt2, vbWide) & "çsñ⁄ [ã@äB/ã@äBáÇ]Ç™èdï°ÇµÇƒÇ‹Ç∑ÅB            "
                                    .MB_Title = ""
                                    .MB_Button = Error
                                    .MBOX
                                End With

                                GoTo Ukeir_Err
                            End If
                        End If
                    Next gCnt2
                End If
            Next gCnt1
        End With
    End If
    '=============================================== *** 2009/07/15 End
    
    
    
    ' ó™èÃè»ó™éûÅAê≥éÆñºèÃÇÊÇËÉZÉbÉg
    If Trim(iB1_Ryaku) = "" Then
        iB1_Ryaku = MidB(iB1_Tornm, 1, 20)
    End If
    
    Body1Chk = True
    Exit Function
Ukeir_Err:
    vsGrid.Row = gLong
    Call TextToVsGrid(Val(iB1_Ukeir.Tag))
    Call VsGridToText(vsGrid.Row)
    gObj.SetFocus
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
    
    If Not TRMRead(gSL_Select, 1) Then
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
        iH1_Torcd = RTrim(TRM.éÊà¯êÊ)
        iH1_Kojyo = RTrim(TRM.çHèÍCD)
        If TRM.éÊà¯êÊãÊï™ = 0 Then
           H1Op_Torkb(0).Value = True
           H1Op_Torkb(0).Tag = 0
        Else
           H1Op_Torkb(1).Value = True
           H1Op_Torkb(0).Tag = 1
        End If
        Call BodyEnbled
        Call BodySet("")
        PB_Kotei.Visible = False
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

    With TRM
        B1Chk_Torkb(0) = .å_ñÒêÊãÊï™
        B1Chk_Torkb(1) = .é–ì‡ãÊï™
        
        iB1_Tornm = RTrim(.ê≥éÆñºèÃ)
        
        '=========================================== 2008/12/26 Start
        iB1_ConvTorcd = RTrim(.ïœä∑å≥éÊà¯êÊ)
        iB1_ConvTorcdKj = RTrim(.ïœä∑å≥éÊà¯êÊçHèÍ)
        '=========================================== 2008/12/26 End
        
        iB1_SKana = RTrim(.çıà¯ÉJÉi)
        iB1_Ryaku = RTrim(.ó™èÃ)
        Call iB1_Ryaku_LostFocus
        
        B1Chk_Taisyo = .â∫êøñ@ëŒè€
        iB1_KojyoNm = RTrim(.çHèÍñº)
        iB1_HTorcd = RTrim(.îÌé©é–)
        iB1_HKojyo = RTrim(.îÌé©é–çHèÍ)
        iB1_KTorcd = RTrim(.å_ñÒêÊ)
'        Call iB1_KTorcd_LostFocus
        B1Op_Keisan(.åvéZï˚ñ@).Value = True
        B1Op_Marume(.åvéZíPà ).Value = True
        
'----------------------------------------------------------***** 2008/06/11 Start
        B1Op_Kurai1(.åvéZà 1).Value = True
        B1Op_Kurai2(.åvéZà 2).Value = True
'----------------------------------------------------------***** 2008/06/11 End
        
        iB1_Shime = CisFun.RSetFld(.í˜ì˙, 0, iB1_Shime.cFormat)
        B1Op_Hei(.ïΩèÄâªãÊï™).Value = True
        iB1_Keisu = CisFun.RSetFld(.ïΩèÄâªåWêî, 0, iB1_Keisu.cFormat)
        iB1_DenSyu = RTrim(.ì`ï[éÌóﬁ)
        Call iB1_DenSyu_LostFocus
'''2007/12/11DEL        CisComb3.Code = RTrim(.î[ïièëå`ë‘)
'''2007/12/11DEL        CisComb3.TextGet
'''2007/12/11DEL        CisComb1.Code = RTrim(.éÛóÃèëå`ë‘)
'''2007/12/11DEL        CisComb1.TextGet
        CisComb2.Code = Format(.î≠íçì‡é¶, "0")
        CisComb2.TextGet
        CisComb4.Code = RTrim(.ñæç◊èëî≠çs)
        CisComb4.TextGet
        
        iB1_Mae = CisFun.RSetFld(.ëOì|ì˙êî, 0, iB1_Mae.cFormat)
        
        B1Chk_PrtkbN = .î[ïièëî≠çsãÊï™
        B1Chk_PrtkbS = .éxããèëî≠çsãÊï™
        B1Chk_PrtkbJ = .éÛóÃèëî≠çsãÊï™
        
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
        If .çÜï‚ãÊï™ = "G" Then
           B1Chk_Gouho.Value = 0
           B1Chk_Gouho.Caption = "çÜå˚"
        Else
           B1Chk_Gouho.Value = 1
           B1Chk_Gouho.Caption = "ï‚ãã"
        End If
        
        B1Chk_MaeHak.Value = .ëOçHíˆãÊï™
        If .ëOçHíˆãÊï™ = 0 Then
           B1Chk_MaeHak.Caption = "ï\é¶ñ≥"
        Else
           B1Chk_MaeHak.Caption = "ï\é¶óL"
        End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
        
        iB1_STHYmd = RTrim(.äJénî≠íçì˙)
        B1Chk_Juryo = .éÛóÃèëî≠çs
        
        iB1_JYubin1 = Mid(.óXï÷î‘çÜ, 1, 3)
        iB1_JYubin2 = Mid(.óXï÷î‘çÜ, 4, 4)
        iB1_JAdd1 = RTrim(.èZèä1)
        iB1_JAdd2 = RTrim(.èZèä2)
        iB1_JTel = RTrim(.ë„ï\TEL)
        iB1_JFax = RTrim(.ë„ï\FAX)
        iB1_JBusyo1 = RTrim(.ïîèêñº1)
        iB1_JBusyo2 = RTrim(.ïîèêñº2)
        iB1_JTTel1 = RTrim(.TEL1)
        iB1_JTTel2 = RTrim(.TEL2)
        iB1_JTFax = RTrim(.FAX)
        iB1_JTEMail = RTrim(.EMAIL)
        iB1_JTanto = RTrim(.íSìñé“)
        iB1_JYaku = RTrim(.ñêEñº)
        iB1_JGinko = RTrim(.ã‚çs)
        iB1_JShiten = RTrim(.éxìX)
        iB1_JGinnm = RTrim(.ã‚çsñº)
        iB1_JHnm = RTrim(.éxìXñº)
        iB1_JKouza = RTrim(.å˚ç¿î‘çÜ)
        iB1_JYokin = Format(.óaã‡ãÊï™, iB1_JYokin.cFormat)
        Call iB1_JYokin_LostFocus
        iB1_JMeigi = RTrim(.ñºã`êl)
        
        iB1_SendFax(0) = .FAXéÊà¯êÊ1
        iB1_SendFax(1) = .FAXéÊà¯êÊ2
        iB1_SendFax(2) = .FAXéÊà¯êÊ3
        iB1_SendMail(0) = .MAILéÊà¯êÊ1
        iB1_SendMail(1) = .MAILéÊà¯êÊ2
        iB1_SendMail(2) = .MAILéÊà¯êÊ3
    
        '=================================== *** 2009/07/06 Start
        'êVe-Ç©ÇÒÇŒÇÒëŒâû
        If H1Op_Torkb(0) Then       'éÛíçÇÃÇ›
           Back_Maekotei(0).Visible = False
           Back_NJigyo.Visible = True
           iB1_NJigyoCD = RTrim(.î[ì¸êÊéñã∆ëÃ)
           iB1_NJigyoKJ = RTrim(.î[ì¸êÊéñã∆ëÃçHãÊ)
           iB1_Syukaba = RTrim(.èoâ◊èÍ)
           If B1Chk_Torkb(0).Value = 1 Then     'å_ñÒêÊÇÃÇ›
              iB1_SJigyoCD = RTrim(.édì¸êÊéñã∆ëÃ)
              Back_Maekotei(1).Visible = False
              Back_SJigyo.Visible = True
              Back_Syukaba.Visible = True
           Else
              Back_Maekotei(1).Visible = True
              Back_SJigyo.Visible = False
              Back_Syukaba.Visible = True
           End If
        Else
           Back_Maekotei(0).Visible = True
           Back_Maekotei(1).Visible = False
           Back_NJigyo.Visible = False
           Back_SJigyo.Visible = False
           Back_Syukaba.Visible = False
        End If
        '=================================== *** 2009/07/06 End
    
        iB1_HinKeta = Format(.ïiî‘É`ÉFÉbÉNåÖêî, "##")           ' 2009.11.06 add
    
        '=================================== *** 2009/07/15 Start
        CisComb5.Code = RTrim(.ïîñÂCD1)
        CisComb5.TextGet
        CisComb6.Code = RTrim(.ïîñÂCD2)
        CisComb6.TextGet
        iB1_SHito = .ê≥é–àıêlàı
        iB1_STime = .ê≥é–àıâ“ì≠éûä‘
        iB1_SGenka = .ê≥é–àıå¥âø
        iB1_PHito = .ÉpÅ[Égêlàı
        iB1_PTime = .ÉpÅ[Égâ“ì≠éûä‘
        iB1_PGenka = .ÉpÅ[Égå¥âø
        iB1_THito = .ÇªÇÃëºêlàı
        iB1_TTime = .ÇªÇÃëºâ“ì≠éûä‘
        iB1_TGenka = .ÇªÇÃëºå¥âø
        B1Chk_Nipo = .ì˙ïÒèoóÕ
        B1Chk_Fuka = .ïââ◊ä«óù
        B1Chk_PSChk = .ÉvÉåÉXÉXÉ|ÉbÉgãÊï™
        '=================================== *** 2009/07/15 Start
        
    End With
        
    Call iB1_KTorcd_LostFocus
    Call iB1_SendFax_LostFocus(0)
    Call iB1_SendFax_LostFocus(1)
    Call iB1_SendFax_LostFocus(2)
    Call iB1_SendMail_LostFocus(0)
    Call iB1_SendMail_LostFocus(1)
    Call iB1_SendMail_LostFocus(2)
    
    ' éÊà¯êÊéÛì¸èÓïÒ
    gSL_Select = "Select * from éÊà¯êÊéÛì¸É}ÉXÉ^ "
    gSL_Select = gSL_Select & " Where éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & RTrim(iH1_Kojyo) & "'"
    gSL_Select = gSL_Select & "   And éÊà¯êÊãÊï™ = " & CisFun.Val2(H1Op_Torkb(0).Tag)
    If TUMRead(gSL_Select, 0, 1) Then
       Call GridSet
    End If
    Call TUMClose(1)
    
    '=================================== *** 2009/07/15 Start
    ' éÊà¯êÊçHíˆèÓïÒ
    If CisFun.Val2(H1Op_Torkb(0).Tag) = 1 And B1Chk_Torkb(1) = 1 Then
       gSL_Select = "Select TK.*,SY1.ílñºèÃ ã@äBñº,SY2.ílñºèÃ ã@äBNOñº,SY3.ílñºèÃ ë„ë÷ã@äBñº,SY4.ílñºèÃ ë„ë÷ã@äBNOñº "
       gSL_Select = gSL_Select & " from éÊà¯êÊçHíˆÉ}ÉXÉ^ TK "
       gSL_Select = gSL_Select & " Left Outer join ñºèÃÉ}ÉXÉ^ SY1"
       gSL_Select = gSL_Select & " On  SY1.ãÊï™ñºèÃ = 'ã@äB' "
       gSL_Select = gSL_Select & " And SY1.íl = TK.ã@äB"
       gSL_Select = gSL_Select & " And SY1.ãÊï™É^ÉCÉv = ''"
       gSL_Select = gSL_Select & " Left Outer join ñºèÃÉ}ÉXÉ^ SY2"
       gSL_Select = gSL_Select & " On  SY2.ãÊï™ñºèÃ = 'ã@äBNO' "
       gSL_Select = gSL_Select & " And SY2.íl = TK.ã@äBNO"
       gSL_Select = gSL_Select & " And SY2.ãÊï™É^ÉCÉv = ''"
       gSL_Select = gSL_Select & " Left Outer join ñºèÃÉ}ÉXÉ^ SY3"
       gSL_Select = gSL_Select & " On  SY3.ãÊï™ñºèÃ = 'ã@äBNO' "
       gSL_Select = gSL_Select & " And SY3.íl = TK.ë„ë÷ã@äB"
       gSL_Select = gSL_Select & " And SY3.ãÊï™É^ÉCÉv = ''"
       gSL_Select = gSL_Select & " Left Outer join ñºèÃÉ}ÉXÉ^ SY4"
       gSL_Select = gSL_Select & " On  SY4.ãÊï™ñºèÃ = 'ã@äBNO' "
       gSL_Select = gSL_Select & " And SY4.íl = TK.ë„ë÷ã@äBNO"
       gSL_Select = gSL_Select & " And SY4.ãÊï™É^ÉCÉv = ''"
       gSL_Select = gSL_Select & " Where éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & RTrim(iH1_Kojyo) & "'"
       If TKTRead(gSL_Select, 0, 1) Then
          Call GridSet2
       End If
    End If
    Call TKTClose(1)
    '=================================== *** 2009/07/15 End
    
    Call BodyEnbled
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
        ' éÊà¯êÊÉ}ÉXÉ^
        .SQL = "DELETE FROM éÊà¯êÊÉ}ÉXÉ^ "
        .SQL = .SQL & " Where éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & RTrim(iH1_Kojyo) & "'"
        .SQL = .SQL & "   And éÊà¯êÊãÊï™ = " & CisFun.Val2(H1Op_Torkb(0).Tag)
        .DBExec
        ' éÊà¯êÊéÛì¸É}ÉXÉ^
        .SQL = "DELETE FROM éÊà¯êÊéÛì¸É}ÉXÉ^ "
        .SQL = .SQL & " Where éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & RTrim(iH1_Kojyo) & "'"
        .SQL = .SQL & "   And éÊà¯êÊãÊï™ = " & CisFun.Val2(H1Op_Torkb(0).Tag)
        On Error GoTo DBDelete_Err
        .DBExec
        
        '========================================================================= 2009/01/06 Start
        ' éÛíçéÊà¯êÊCDïœä∑É}ÉXÉ^
        If H1Op_Torkb(0) Then
           .SQL = "DELETE FROM éÛíçéÊà¯êÊCDïœä∑É}ÉXÉ^ "
           .SQL = .SQL & " Where ïœä∑êÊCD = '" & RTrim(iH1_Torcd) & RTrim(iH1_Kojyo) & "'"
           On Error GoTo DBDelete_Err
           .DBExec
        End If
        '========================================================================= 2009/01/06 End
        
        '========================================================================= 2009/07/15 Start
        ' éÊà¯êÊçHíˆÉ}ÉXÉ^
        If H1Op_Torkb(1) Then
           .SQL = "DELETE FROM éÊà¯êÊçHíˆÉ}ÉXÉ^ "
           .SQL = .SQL & " Where éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & RTrim(iH1_Kojyo) & "'"
           On Error GoTo DBDelete_Err
           .DBExec
        End If
        '========================================================================= 2009/07/15 End
        
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
    Dim AddFlg      As Boolean
    Dim wDenpyo     As String
    Dim wNouKei     As String
    Dim wJyuKei     As String
    Dim wJyuPkb     As Byte
    Dim wPrtkbN     As Byte
    Dim wPrtkbS     As Byte
    Dim wPrtkbJ     As Byte
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
    Dim wGouho      As String
    Dim wMaeHak     As Byte
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
    Dim wMeiPkb     As String
    Dim wWhere1     As String
    Dim wWhere2     As String
    Dim wWhere3     As String
    
    '============================== *** 2009/07/06 Start
    Dim wSyukaba  As String
    '============================== *** 2009/07/06 End
    
    '============================== *** 2009/07/15 Start
    Dim wFuka       As String
    '============================== *** 2009/07/15 End
    
    Dim wUkeir()    As String
    Dim wCycle()    As String
    Dim wwUkeir()   As String
    Dim wwCycle()   As String
    Dim wYoubi(5)   As Byte

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

'   É}ÉXÉ^çÄñ⁄à⁄ëó
    wDenpyo = "": wNouKei = "": wJyuKei = "": wJyuPkb = 0: wMeiPkb = ""
    gSL_Select = "SELECT * From éÊà¯êÊÉ}ÉXÉ^"
    gSL_Select = gSL_Select & " Where éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & RTrim(iH1_Kojyo) & "'"
    gSL_Select = gSL_Select & "   And éÊà¯êÊãÊï™ = " & CisFun.Val2(H1Op_Torkb(0).Tag)
    If TRMRead(gSL_Select, 1) Then
        AddFlg = False
        'ïœçXëOëﬁî
        wDenpyo = TRM.ì`ï[éÌóﬁ
        wNouKei = TRM.î[ïièëå`ë‘
        wJyuKei = TRM.éÛóÃèëå`ë‘
        wJyuPkb = TRM.éÛóÃèëî≠çs
        wMeiPkb = TRM.ñæç◊èëî≠çs
        wPrtkbN = TRM.î[ïièëî≠çsãÊï™
        wPrtkbS = TRM.éxããèëî≠çsãÊï™
        wPrtkbJ = TRM.éÛóÃèëî≠çsãÊï™
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
        wGouho = TRM.çÜï‚ãÊï™
        wMaeHak = TRM.ëOçHíˆãÊï™
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
        '================================= *** 2009/07/06 Start
        wSyukaba = TRM.èoâ◊èÍ
        '================================= *** 2009/07/06 End
        
        '================================= *** 2009/07/15 Start
        wFuka = TRM.ïââ◊ä«óù
        '================================= *** 2009/07/15 End
    Else
        Call ItemsClearTRM
        AddFlg = True
        TRM.éÊà¯êÊCD = RTrim(iH1_Torcd) & RTrim(iH1_Kojyo)
        TRM.éÊà¯êÊãÊï™ = CisFun.Val2(H1Op_Torkb(0).Tag)
        TRM.éÊà¯êÊ = RTrim(iH1_Torcd)
        TRM.çHèÍCD = RTrim(iH1_Kojyo)
    End If
    With TRM
        .å_ñÒêÊãÊï™ = B1Chk_Torkb(0).Value
        .é–ì‡ãÊï™ = B1Chk_Torkb(1).Value
        .ê≥éÆñºèÃ = RTrim(iB1_Tornm)
        '====================================== 2008/12/26 Start
        .ïœä∑å≥éÊà¯êÊ = RTrim(iB1_ConvTorcd)
        .ïœä∑å≥éÊà¯êÊçHèÍ = RTrim(iB1_ConvTorcdKj)
        .ïœä∑å≥éÊà¯êÊCD = RTrim(iB1_ConvTorcd) & RTrim(iB1_ConvTorcdKj)
        '====================================== 2008/12/26 End
        .çıà¯ÉJÉi = RTrim(iB1_SKana)
        .ó™èÃ = RTrim(iB1_Ryaku)
        If B1Chk_Taisyo.Visible Then
           .â∫êøñ@ëŒè€ = B1Chk_Taisyo.Value
        Else
           .â∫êøñ@ëŒè€ = 0
        End If
        If iB1_KojyoNm.Visible Then
           .çHèÍñº = RTrim(iB1_KojyoNm)
        Else
           .çHèÍñº = ""
        End If
        If iB1_KTorcd.Visible Then
           .å_ñÒêÊ = RTrim(iB1_KTorcd)
        Else
           .å_ñÒêÊ = ""
        End If
        If iB1_HTorcd.Visible Then
            .îÌé©é– = iB1_HTorcd
            .îÌé©é–çHèÍ = iB1_HKojyo
            .îÌé©é–CD = RTrim(iB1_HTorcd) & RTrim(iB1_HKojyo)
        Else
            .îÌé©é– = ""
            .îÌé©é–çHèÍ = ""
            .îÌé©é–CD = ""
        End If
        If Back_KeisanH.Visible Then
            .åvéZï˚ñ@ = CisFun.Val2(B1Op_Keisan(0).Tag)
            .åvéZíPà  = CisFun.Val2(B1Op_Marume(0).Tag)
            '------------------------------------------------------***** 2008/06/11 Start
            If .åvéZíPà  = 3 Then
               .åvéZà 1 = CisFun.Val2(B1Op_Kurai1(0).Tag)
               .åvéZà 2 = CisFun.Val2(B1Op_Kurai2(0).Tag)
            Else
               .åvéZà 1 = 0
               .åvéZà 2 = 0
            End If
            '------------------------------------------------------***** 2008/06/11 End
        Else
            .åvéZï˚ñ@ = 0
            .åvéZíPà  = 0
            '------------------------------------------------------***** 2008/06/11 Start
            .åvéZà 1 = 0
            .åvéZà 2 = 0
            '------------------------------------------------------***** 2008/06/11 End
        End If
        If Back_Heijyun.Visible Then
           .ïΩèÄâªãÊï™ = CisFun.Val2(B1Op_Hei(0).Tag)
        Else
           .ïΩèÄâªãÊï™ = 0
        End If
        If iB1_Shime.Visible Then
           .í˜ì˙ = CisFun.Val2(iB1_Shime)
        Else
           .í˜ì˙ = 0
        End If
        If iB1_Keisu.Visible Then
           .ïΩèÄâªåWêî = CisFun.Val2(iB1_Keisu)
        Else
           .ïΩèÄâªåWêî = 0
        End If
        If iB1_DenSyu.Visible Then
            .ì`ï[éÌóﬁ = RTrim(iB1_DenSyu)
            .î[ïièëå`ë‘ = SaveKeitai
            .éÛóÃèëå`ë‘ = SaveKeitai
        Else
           .ì`ï[éÌóﬁ = ""
           .î[ïièëå`ë‘ = ""
           .éÛóÃèëå`ë‘ = ""
        End If
'''2007/12/11DEL        If B1Comb_Nouhin.Visible Then
'''2007/12/11DEL            CisComb3.CodeGet
'''2007/12/11DEL            .î[ïièëå`ë‘ = RTrim(CisComb3.Code)
'''2007/12/11DEL        Else
'''2007/12/11DEL            .î[ïièëå`ë‘ = ""
'''2007/12/11DEL        End If
    '-------
        If B1Chk_PrtkbN.Visible Then
            .î[ïièëî≠çsãÊï™ = B1Chk_PrtkbN.Value
        Else
            .î[ïièëî≠çsãÊï™ = 0
        End If
        If B1Chk_PrtkbS.Visible Then
            .éxããèëî≠çsãÊï™ = B1Chk_PrtkbS.Value
        Else
            .éxããèëî≠çsãÊï™ = 0
        End If
        If B1Chk_PrtkbJ.Visible Then
            .éÛóÃèëî≠çsãÊï™ = B1Chk_PrtkbJ.Value
        Else
            .éÛóÃèëî≠çsãÊï™ = 0
        End If
    '-------
        
'''2007/12/11DEL        If B1Comb_Jyuryo.Visible Then
'''2007/12/11DEL            CisComb1.CodeGet
'''2007/12/11DEL            .éÛóÃèëå`ë‘ = RTrim(CisComb1.Code)
'''2007/12/11DEL        Else
'''2007/12/11DEL            .éÛóÃèëå`ë‘ = ""
'''2007/12/11DEL        End If
        If B1Chk_Juryo.Visible Then
            .éÛóÃèëî≠çs = B1Chk_Juryo.Value
        Else
            .éÛóÃèëî≠çs = 0
        End If
        If B1Comb_HNaiji.Visible Then
            CisComb2.CodeGet
            .î≠íçì‡é¶ = CisFun.Val2(CisComb2.Code)
        Else
            .î≠íçì‡é¶ = 0
        End If
        If B1Comb_Meaisai.Visible Then
            CisComb4.CodeGet
            .ñæç◊èëî≠çs = RTrim(CisComb4.Code)
        Else
            .ñæç◊èëî≠çs = ""
        End If
        If iB1_Mae.Visible Then
           .ëOì|ì˙êî = CisFun.Val2(iB1_Mae)
        Else
           .ëOì|ì˙êî = 0
        End If
        .äJénî≠íçì˙ = RTrim(iB1_STHYmd)
        .óXï÷î‘çÜ = RTrim(iB1_JYubin1) & RTrim(iB1_JYubin2)
        .èZèä1 = RTrim(iB1_JAdd1)
        .èZèä2 = RTrim(iB1_JAdd2)
        .ë„ï\TEL = RTrim(iB1_JTel)
        .ë„ï\FAX = RTrim(iB1_JFax)
        .ïîèêñº1 = RTrim(iB1_JBusyo1)
        .ïîèêñº2 = RTrim(iB1_JBusyo2)
        .TEL1 = RTrim(iB1_JTTel1)
        .TEL2 = RTrim(iB1_JTTel2)
        .FAX = RTrim(iB1_JTFax)
        .EMAIL = RTrim(iB1_JTEMail)
        .íSìñé“ = RTrim(iB1_JTanto)
        .ñêEñº = RTrim(iB1_JYaku)
        
        If Back_Ginko.Visible Then
            .ã‚çs = RTrim(iB1_JGinko)
            .éxìX = RTrim(iB1_JShiten)
            .ã‚çsñº = RTrim(iB1_JGinnm)
            .éxìXñº = RTrim(iB1_JHnm)
            .å˚ç¿î‘çÜ = RTrim(iB1_JKouza)
            .óaã‡ãÊï™ = CisFun.Val2(iB1_JYokin)
            .ñºã`êl = RTrim(iB1_JMeigi)
        Else
            .ã‚çs = ""
            .éxìX = ""
            .ã‚çsñº = ""
            .éxìXñº = ""
            .å˚ç¿î‘çÜ = ""
            .óaã‡ãÊï™ = 0
            .ñºã`êl = ""
        End If
        
        .FAXéÊà¯êÊ1 = ""
        .FAXéÊà¯êÊ2 = ""
        .FAXéÊà¯êÊ3 = ""
        .MAILéÊà¯êÊ1 = ""
        .MAILéÊà¯êÊ2 = ""
        .MAILéÊà¯êÊ3 = ""
        If Back_Send.Visible Then
           For gCnt1 = 0 To 2
               If Trim(iB1_SendFax(gCnt1)) = "" Then
                  For gCnt2 = gCnt1 + 1 To 2
                      If Trim(iB1_SendFax(gCnt2)) <> "" Then
                         iB1_SendFax(gCnt1) = iB1_SendFax(gCnt2)
                         iB1_SendFax(gCnt2) = ""
                         Exit For
                      End If
                  Next gCnt2
                End If
           Next gCnt1
           For gCnt1 = 0 To 2
               If Trim(iB1_SendMail(gCnt1)) = "" Then
                  For gCnt2 = gCnt1 + 1 To 2
                      If Trim(iB1_SendMail(gCnt2)) <> "" Then
                         iB1_SendMail(gCnt1) = iB1_SendMail(gCnt2)
                         iB1_SendMail(gCnt2) = ""
                         Exit For
                      End If
                  Next gCnt2
                End If
           Next gCnt1
           
           .FAXéÊà¯êÊ1 = iB1_SendFax(0)
           .FAXéÊà¯êÊ2 = iB1_SendFax(1)
           .FAXéÊà¯êÊ3 = iB1_SendFax(2)
           .MAILéÊà¯êÊ1 = iB1_SendMail(0)
           .MAILéÊà¯êÊ2 = iB1_SendMail(1)
           .MAILéÊà¯êÊ3 = iB1_SendMail(2)
        End If
        
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
        If B1Chk_Gouho.Value = 0 Then
           .çÜï‚ãÊï™ = "G"
        Else
           .çÜï‚ãÊï™ = "H"
        End If
        If B1Chk_MaeHak.Visible Then
           .ëOçHíˆãÊï™ = B1Chk_MaeHak.Value
        End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
        
        '============================================== *** 2009/07/06 Start
        .î[ì¸êÊéñã∆ëÃ = RTrim(iB1_NJigyoCD)
        .î[ì¸êÊéñã∆ëÃçHãÊ = RTrim(iB1_NJigyoKJ)
        .édì¸êÊéñã∆ëÃ = RTrim(iB1_SJigyoCD)
        .èoâ◊èÍ = RTrim(iB1_Syukaba)
        '============================================== *** 2009/07/06 End
    
        '============================================== 2009/07/15 Start
        If B1Chk_Torkb(1).Value = 1 Then
            CisComb5.CodeGet
            CisComb6.CodeGet
           .ïîñÂCD1 = RTrim(CisComb5.Code)
           .ïîñÂCD2 = RTrim(CisComb6.Code)
           .ê≥é–àıêlàı = CisFun.Val2(iB1_SHito)
           .ê≥é–àıâ“ì≠éûä‘ = CisFun.Val2(iB1_STime)
           .ê≥é–àıå¥âø = CisFun.Val2(iB1_SGenka)
           .ÉpÅ[Égêlàı = CisFun.Val2(iB1_PHito)
           .ÉpÅ[Égâ“ì≠éûä‘ = CisFun.Val2(iB1_PTime)
           .ÉpÅ[Égå¥âø = CisFun.Val2(iB1_PGenka)
           .ÇªÇÃëºêlàı = CisFun.Val2(iB1_THito)
           .ÇªÇÃëºâ“ì≠éûä‘ = CisFun.Val2(iB1_TTime)
           .ÇªÇÃëºå¥âø = CisFun.Val2(iB1_TGenka)
           .ì˙ïÒèoóÕ = B1Chk_Nipo.Value
           .ïââ◊ä«óù = B1Chk_Fuka.Value
           .ÉvÉåÉXÉXÉ|ÉbÉgãÊï™ = B1Chk_PSChk.Value
        End If
        If H1Op_Torkb(0).Value Then
           .éwé¶èëî≠çsãÊï™ = B1Chk_Sijisyo.Value
        End If
        '============================================== 2009/07/15 End
        
        '---------------------------------------------------------------' 2009.11.06 add start
        If H1Op_Torkb(0).Value = True And Trim(iB1_HinKeta) <> "" Then
           .ïiî‘É`ÉFÉbÉNåÖêî = CisFun.Val2(iB1_HinKeta)
        Else
           .ïiî‘É`ÉFÉbÉNåÖêî = 0
        End If
        '---------------------------------------------------------------' 2009.11.06 add end
    End With

'-----------------------------------------------
'   éÊà¯êÊÉ}ÉXÉ^çXêV
'-----------------------------------------------
    Call CisDB.DBTran(TransBegin)  ' ƒ◊›ªﬁ∏ºÆ›äJén
' ( í«â¡èàóù )
    If AddFlg Then
        TRM.çÏê¨é“ = gTanto
        Call TRMInsert
        iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
    Else
' ( çXêVèàóù )
        TRM.çXêVé“ = gTanto
        Call TRMUpdate
        iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
    End If

'-----------------------------------------------
'   éÊà¯êÊéÛì¸É}ÉXÉ^çXêV
'-----------------------------------------------
    ReDim wUkeir(1)
    ReDim wCycle(1)
    gSL_Select = "select * from éÊà¯êÊéÛì¸É}ÉXÉ^ "
    gSL_Select = gSL_Select & " WHERE éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & RTrim(iH1_Kojyo) & "'"
    gSL_Select = gSL_Select & "   And éÊà¯êÊãÊï™ = " & CisFun.Val2(H1Op_Torkb(0).Tag)
    If TUMRead(gSL_Select) Then
       ReDim wUkeir(CisDB.RecordCount)
       ReDim wCycle(CisDB.RecordCount)
       gCnt1 = 0
       Do Until Not TUM_RDSTS
          wUkeir(gCnt1) = TUM.éÛì¸
          wCycle(gCnt1) = TUM.ÉTÉCÉNÉã
            
          gCnt1 = gCnt1 + 1
          Call TUMReadNext
       Loop
    End If
    Call TUMClose
            
    With CisDB
        .SQL = "DELETE éÊà¯êÊéÛì¸É}ÉXÉ^ "
        .SQL = .SQL & " WHERE éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & RTrim(iH1_Kojyo) & "'"
        .SQL = .SQL & "   And éÊà¯êÊãÊï™ = " & CisFun.Val2(H1Op_Torkb(0).Tag)
        .DBExec
    End With

    With vsGrid
        ReDim wwUkeir(.Rows - 1)
        ReDim wwCycle(.Rows - 1)
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, 1)) <> "" Then
               Call ItemsClearTUM
               With TUM
                    .éÊà¯êÊCD = RTrim(iH1_Torcd) & RTrim(iH1_Kojyo)
                    .éÊà¯êÊãÊï™ = CisFun.Val2(H1Op_Torkb(0).Tag)
                    .éÛì¸ = Trim(vsGrid.TextMatrix(gLong, 1))
                    .éÛì¸ñºèÃ = Trim(vsGrid.TextMatrix(gLong, 2))
                    .ÉTÉCÉNÉã = CycleFormat(CisFun.Val2(gLong), "")
                     wwUkeir(gLong - 1) = .éÛì¸
                     wwCycle(gLong - 1) = .ÉTÉCÉNÉã
                     If Trim(vsGrid.TextMatrix(gLong, 4)) = "ñ≥" Then
                        .ãÛñæç◊ = "1"
                     Else
                        .ãÛñæç◊ = ""
                     End If
                     If Trim(vsGrid.TextMatrix(gLong, 5)) = "Ç∑ÇÈ" Then
                        .ójì˙î≠íç = 1
                     Else
                        .ójì˙î≠íç = 0
                     End If
                     If Trim(vsGrid.TextMatrix(gLong, 6)) = "Åõ" Then
                        .ójì˙1 = 1
                     Else
                        .ójì˙1 = 0
                     End If
                     If Trim(vsGrid.TextMatrix(gLong, 7)) = "Åõ" Then
                        .ójì˙2 = 1
                     Else
                        .ójì˙2 = 0
                     End If
                     If Trim(vsGrid.TextMatrix(gLong, 8)) = "Åõ" Then
                        .ójì˙3 = 1
                     Else
                        .ójì˙3 = 0
                     End If
                     If Trim(vsGrid.TextMatrix(gLong, 9)) = "Åõ" Then
                        .ójì˙4 = 1
                     Else
                        .ójì˙4 = 0
                     End If
                     If Trim(vsGrid.TextMatrix(gLong, 10)) = "Åõ" Then
                        .ójì˙5 = 1
                     Else
                        .ójì˙5 = 0
                     End If
               End With
               Call TUMInsert
            End If
        Next gLong
    End With
    
'========================================================== *** 2009/07/15 Start

    If CisFun.Val2(H1Op_Torkb(0).Tag) = 1 Then
        With CisDB
            .SQL = "DELETE éÊà¯êÊçHíˆÉ}ÉXÉ^ "
            .SQL = .SQL & " WHERE éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & RTrim(iH1_Kojyo) & "'"
            .DBExec
        End With
    
        With VsGrid2
            For gLong = 1 To .Rows - 1
                If Trim(.TextMatrix(gLong, 1)) <> "" Or Trim(.TextMatrix(gLong, 3)) <> "" Then
                   Call ItemsClearTKT
                   With TKT
                        .éÊà¯êÊCD = RTrim(iH1_Torcd) & RTrim(iH1_Kojyo)
                        .ã@äB = Trim(VsGrid2.TextMatrix(gLong, 1))
                        .ã@äBNO = Trim(VsGrid2.TextMatrix(gLong, 3))
                        .â“ì≠éûä‘ = CisFun.Val2(Trim(VsGrid2.TextMatrix(gLong, 5)))
                        .ë‰êî = CisFun.Val2(Trim(VsGrid2.TextMatrix(gLong, 6)))
                        .ë„ë÷ã@äB = Trim(VsGrid2.TextMatrix(gLong, 7))
                        .ë„ë÷ã@äBNO = Trim(VsGrid2.TextMatrix(gLong, 9))
                   End With
                   Call TKTInsert
                End If
            Next gLong
        End With
    End If
    
'========================================================== *** 2009/07/15 End

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
'+------------------------------------------------------+
'+     çÜï‚ãÊï™ÇÃïœçXÇ™Ç†Ç¡ÇΩèÍçá,                      +
'+     ämîFÉÅÉbÉZÅ[ÉWï\é¶å„ÅAïiî‘éÊà¯êÊÉ}ÉXÉ^ÇÃçXêV     +
'+------------------------------------------------------+
    If Not AddFlg And wGouho <> TRM.çÜï‚ãÊï™ Then
       '============= ìØàÍèÓïÒÇéùÇ¬ïiî‘éÊà¯êÊÉ}ÉXÉ^ÇÃämîF =================
       gSL_Select = ""
       gSL_Select = gSL_Select & " SELECT COUNT(ïiî‘) åèêî FROM ïiî‘éÊà¯êÊÉ}ÉXÉ^"
       gSL_Select = gSL_Select & "  WHERE éÊà¯êÊ = '" & RTrim(TRM.éÊà¯êÊCD) & "'"
       gSL_Select = gSL_Select & "    AND ïiñ⁄ = " & RTrim(TRM.éÊà¯êÊãÊï™)
       gSL_Select = gSL_Select & "    AND çÜï‚ãÊï™ = '" & RTrim(TRM.çÜï‚ãÊï™) & "'"
       If HTMRead(gSL_Select, 1) Then
          If HTM.åèêî <> 0 Then
             With CisFun
                  '============== çÜï‚ãÊï™ïœçXÇÃämîFÉÅÉbÉZÅ[ÉW ================
                  .MB_Lines = 6
                  .MB_MSG(1) = "     çÜï‚ãÊï™Ç™ïœçXÇ≥ÇÍÇ‹ÇµÇΩÅB      "
                  .MB_MSG(3) = "     ïœçXëOÇÃílÇéùÇ¬Ç±ÇÃéÊà¯êÊÇÃïiî‘Ç™ " & HTM.åèêî & " åèë∂ç›ÇµÇ‹Ç∑ÅB      "
                  .MB_MSG(5) = "     ïiî‘É}ÉXÉ^ÇÃ" & gStr & "ÇïœçXå„ÇÃílÇ…àÍäáíuä∑ÇµÇ‹Ç∑Ç©ÅH      "
                  .MB_Title = "çÜï‚ãÊï™íuä∑èàóù"
                  .MB_Button = OK_CAN
                  If .MBOX() Then
                     With CisDB
                          '============== ïiî‘éÊà¯êÊÉ}ÉXÉ^çXêV ================
                          .SQL = ""
                          .SQL = .SQL & " UPDATE ïiî‘éÊà¯êÊÉ}ÉXÉ^ SET" & vbCrLf
                          .SQL = .SQL & "        çÜï‚ãÊï™ = '" & RTrim(TRM.çÜï‚ãÊï™) & "'," & vbCrLf
                          .SQL = .SQL & "        çXêVì˙  = GETDATE()," & vbCrLf
                          .SQL = .SQL & "        çXêVí[ññ  = HOST_NAME()," & vbCrLf
                          .SQL = .SQL & "        çXêVé“  = '" & gTanto & "'" & vbCrLf
                          .SQL = .SQL & "  WHERE éÊà¯êÊ = '" & RTrim(TRM.éÊà¯êÊ) & "'" & vbCrLf
                          .SQL = .SQL & "    AND ïiñ⁄ = " & RTrim(TRM.éÊà¯êÊãÊï™) & vbCrLf
                          .SQL = .SQL & "    AND çÜï‚ãÊï™ = '" & RTrim(TRM.çÜï‚ãÊï™) & "'" & vbCrLf
                          
                          On Error GoTo DBPut_Err
                          .DBExec
                     End With
                  End If
             End With
          End If
       End If
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
    
'-------------------------------------------------------------------------------------------'
'   ì`ï[éÌóﬁÅEî[ïièëå`ë‘ÅEéÛóÃèëå`ë‘ÅEéÛóÃèëî≠çsÅEñæç◊èëî≠çsÇ™ïœçXÇ≥ÇÍÇΩèÍçáÅA
'       ämîFå„ÅAïiî‘éÊà¯êÊÉ}ÉXÉ^ÇçXêV
'-------------------------------------------------------------------------------------------'
    If Not AddFlg Then
        If wDenpyo <> TRM.ì`ï[éÌóﬁ Or wNouKei <> TRM.î[ïièëå`ë‘ Or wJyuKei <> TRM.éÛóÃèëå`ë‘ Or _
            wJyuPkb <> TRM.éÛóÃèëî≠çs Or wMeiPkb <> TRM.ñæç◊èëî≠çs Or _
            wPrtkbN <> TRM.î[ïièëî≠çsãÊï™ Or wPrtkbS <> TRM.éxããèëî≠çsãÊï™ Or wPrtkbJ <> TRM.éÛóÃèëî≠çsãÊï™ Then
            wWhere1 = " WHERE éÊà¯êÊ = '" & RTrim(TRM.éÊà¯êÊCD) & "'"
            wWhere2 = "   AND ïiñ⁄ = " & RTrim(TRM.éÊà¯êÊãÊï™)
            gCnt1 = 0: gStr = "": wWhere3 = ""
            If wDenpyo <> TRM.ì`ï[éÌóﬁ Then
                wWhere3 = wWhere3 & "   AND (ì`ï[éÌóﬁ = '" & RTrim(wDenpyo) & "'"
                gStr = gStr & "( ì`ï[éÌóﬁ "
                gCnt1 = gCnt1 + 1
            End If
            If wNouKei <> TRM.î[ïièëå`ë‘ Then
                If gCnt1 = 0 Then
                    wWhere3 = wWhere3 & " AND ("
                    gStr = gStr & "( î[ïièëå`ë‘ "
                Else
                    wWhere3 = wWhere3 & " OR "
                    gStr = gStr & "• î[ïièëå`ë‘ "
                End If
                wWhere3 = wWhere3 & "  î[ïièëå`ë‘ = '" & RTrim(wNouKei) & "'"
                gCnt1 = gCnt1 + 1
            End If
            If wJyuKei <> TRM.éÛóÃèëå`ë‘ Then
                If gCnt1 = 0 Then
                    wWhere3 = wWhere3 & " AND ("
                    gStr = gStr & "( éÛóÃèëå`ë‘ "
                Else
                    wWhere3 = wWhere3 & " OR "
                    gStr = gStr & "• éÛóÃèëå`ë‘ "
                End If
                wWhere3 = wWhere3 & "  éÛóÃèëå`ë‘ = '" & RTrim(wJyuKei) & "'"
                gCnt1 = gCnt1 + 1
            End If
            If wJyuPkb <> TRM.éÛóÃèëî≠çs Then
                If gCnt1 = 0 Then
                    wWhere3 = wWhere3 & " AND ("
                    gStr = gStr & "( éÛóÃèëî≠çs "
                Else
                    wWhere3 = wWhere3 & " OR "
                    gStr = gStr & "• éÛóÃèëî≠çs "
                End If
                wWhere3 = wWhere3 & "  éÛóÃèëî≠çs = " & wJyuPkb & ""
                gCnt1 = gCnt1 + 1
            End If
            If wMeiPkb <> TRM.ñæç◊èëî≠çs Then
                If gCnt1 = 0 Then
                    wWhere3 = wWhere3 & " AND ("
                    gStr = gStr & "( ñæç◊èëî≠çs "
                Else
                    wWhere3 = wWhere3 & " OR "
                    gStr = gStr & "• ñæç◊èëî≠çs "
                End If
                wWhere3 = wWhere3 & "  ñæç◊èëî≠çs = '" & RTrim(wMeiPkb) & "'"
                gCnt1 = gCnt1 + 1
            End If
            If wPrtkbN <> TRM.î[ïièëî≠çsãÊï™ Then
                If gCnt1 = 0 Then
                    wWhere3 = wWhere3 & " AND ("
                    gStr = gStr & "( î[ïièëî≠çsãÊï™ "
                Else
                    wWhere3 = wWhere3 & " OR "
                    gStr = gStr & "• î[ïièëî≠çsãÊï™ "
                End If
                wWhere3 = wWhere3 & "  î[ïièëî≠çsãÊï™ = " & wPrtkbN & ""
                gCnt1 = gCnt1 + 1
            End If
            If wPrtkbS <> TRM.éxããèëî≠çsãÊï™ Then
                If gCnt1 = 0 Then
                    wWhere3 = wWhere3 & " AND ("
                    gStr = gStr & "( éxããèëî≠çsãÊï™ "
                Else
                    wWhere3 = wWhere3 & " OR "
                    gStr = gStr & "• éxããèëî≠çsãÊï™ "
                End If
                wWhere3 = wWhere3 & "  éxããèëî≠çsãÊï™ = " & wPrtkbS & ""
                gCnt1 = gCnt1 + 1
            End If
            If wPrtkbJ <> TRM.éÛóÃèëî≠çsãÊï™ Then
                If gCnt1 = 0 Then
                    wWhere3 = wWhere3 & " AND ("
                    gStr = gStr & "( éÛóÃèëî≠çsãÊï™ "
                Else
                    wWhere3 = wWhere3 & " OR "
                    gStr = gStr & "• éÛóÃèëî≠çsãÊï™ "
                End If
                wWhere3 = wWhere3 & "  éÛóÃèëî≠çsãÊï™ = " & RTrim(wPrtkbJ) & ""
                gCnt1 = gCnt1 + 1
            End If
            If gCnt1 <> 0 Then
                wWhere3 = wWhere3 & ")"
                gStr = gStr & ")"
            End If
            '(( ïiî‘éÊà¯êÊÉ}ÉXÉ^çXêV ))
            gSL_Select = "SELECT COUNT(ïiî‘) åèêî FROM ïiî‘éÊà¯êÊÉ}ÉXÉ^"
            gSL_Select = gSL_Select & wWhere1 & wWhere2 & wWhere3
            If HTMRead(gSL_Select, 1) Then
                If HTM.åèêî <> 0 Then
                    With CisFun
                        .MB_Lines = 6
                        .MB_MSG(1) = "     " & gStr & " Ç™ïœçXÇ≥ÇÍÇ‹ÇµÇΩÅB      "
                        .MB_MSG(3) = "     ïœçXëOÇÃílÇéùÇ¬Ç±ÇÃéÊà¯êÊÇÃïiî‘Ç™ " & HTM.åèêî & " åèë∂ç›ÇµÇ‹Ç∑ÅB      "
                        .MB_MSG(5) = "     ïiî‘É}ÉXÉ^ÇÃ" & gStr & "ÇïœçXå„ÇÃílÇ…àÍäáíuä∑ÇµÇ‹Ç∑Ç©ÅH      "
                        .MB_Title = "ì`ï[çÄñ⁄íuä∑èàóù"
                        .MB_Button = OK_CAN
                        If .MBOX Then
                            '-------------------------------------------------'
                            '   ämîFå„ÅAïiî‘éÊà¯êÊÉ}ÉXÉ^ÇçXêV
                            '-------------------------------------------------'
                            gSL_Select = "UPDATE ïiî‘éÊà¯êÊÉ}ÉXÉ^ SET "
                            gSL_Select = gSL_Select & "ì`ï[éÌóﬁ  = CASE WHEN ì`ï[éÌóﬁ = '" & RTrim(wDenpyo) & "' THEN '"
                            gSL_Select = gSL_Select & RTrim(TRM.ì`ï[éÌóﬁ) & "' ELSE ì`ï[éÌóﬁ END ,"
                            gSL_Select = gSL_Select & "î[ïièëå`ë‘  = CASE WHEN î[ïièëå`ë‘ = '" & RTrim(wNouKei) & "' THEN '"
                            gSL_Select = gSL_Select & RTrim(TRM.î[ïièëå`ë‘) & "' ELSE î[ïièëå`ë‘ END ,"
                            gSL_Select = gSL_Select & "éÛóÃèëå`ë‘  = CASE WHEN éÛóÃèëå`ë‘ = '" & RTrim(wJyuKei) & "' THEN '"
                            gSL_Select = gSL_Select & RTrim(TRM.éÛóÃèëå`ë‘) & "' ELSE éÛóÃèëå`ë‘ END ,"
                            gSL_Select = gSL_Select & "éÛóÃèëî≠çs  = CASE WHEN éÛóÃèëî≠çs = " & wJyuPkb & " THEN "
                            gSL_Select = gSL_Select & TRM.éÛóÃèëî≠çs & " ELSE éÛóÃèëî≠çs END ,"
                            gSL_Select = gSL_Select & "ñæç◊èëî≠çs  = CASE WHEN ñæç◊èëî≠çs = '" & RTrim(wMeiPkb) & "' THEN '"
                            gSL_Select = gSL_Select & RTrim(TRM.ñæç◊èëî≠çs) & "' ELSE ñæç◊èëî≠çs END ,"
                            gSL_Select = gSL_Select & "î[ïièëî≠çsãÊï™  = CASE WHEN î[ïièëî≠çsãÊï™ = " & wPrtkbN & " THEN "
                            gSL_Select = gSL_Select & TRM.î[ïièëî≠çsãÊï™ & " ELSE î[ïièëî≠çsãÊï™ END ,"
                            gSL_Select = gSL_Select & "éxããèëî≠çsãÊï™  = CASE WHEN éxããèëî≠çsãÊï™ = " & wPrtkbS & " THEN "
                            gSL_Select = gSL_Select & TRM.éxããèëî≠çsãÊï™ & " ELSE éxããèëî≠çsãÊï™ END ,"
                            gSL_Select = gSL_Select & "éÛóÃèëî≠çsãÊï™  = CASE WHEN éÛóÃèëî≠çsãÊï™ = " & wPrtkbJ & " THEN "
                            gSL_Select = gSL_Select & TRM.éÛóÃèëî≠çsãÊï™ & " ELSE éÛóÃèëî≠çsãÊï™ END ,"
                            gSL_Select = gSL_Select & "çXêVì˙  = GETDATE(),"
                            gSL_Select = gSL_Select & "çXêVí[ññ  = HOST_NAME(),"
                            gSL_Select = gSL_Select & "çXêVé“  = '" & gTanto & "'"
                            gSL_Select = gSL_Select & wWhere1 & wWhere2 & wWhere3
                            With CisDB
                                .SQL = gSL_Select
                                On Error GoTo DBPut_Err
                                .DBExec
                            End With
                        End If
                    End With
                End If
            End If
            If TRM.éÊà¯êÊãÊï™ = 1 Then      'î≠íçÇÃéû
                '(( çﬁóøéÊà¯êÊÉ}ÉXÉ^çXêV ))
                gSL_Select = "SELECT COUNT(çﬁóøä«óùî‘çÜ) åèêî FROM çﬁóøéÊà¯êÊÉ}ÉXÉ^"
                gSL_Select = gSL_Select & wWhere1 & wWhere3
                If HTMRead(gSL_Select, 1) Then
                    If HTM.åèêî <> 0 Then
                        With CisFun
                            .MB_Lines = 6
                            .MB_MSG(1) = "     " & gStr & " Ç™ïœçXÇ≥ÇÍÇ‹ÇµÇΩÅB      "
                            .MB_MSG(3) = "     ïœçXëOÇÃílÇéùÇ¬Ç±ÇÃéÊà¯êÊÇÃçﬁóøÇ™ " & HTM.åèêî & " åèë∂ç›ÇµÇ‹Ç∑ÅB      "
                            .MB_MSG(5) = "     çﬁóøÉ}ÉXÉ^ÇÃ" & gStr & "ÇïœçXå„ÇÃílÇ…àÍäáíuä∑ÇµÇ‹Ç∑Ç©ÅH      "
                            .MB_Title = "ì`ï[çÄñ⁄íuä∑èàóù"
                            .MB_Button = OK_CAN
                            If .MBOX Then
                                '-------------------------------------------------'
                                '   ämîFå„ÅAïiî‘éÊà¯êÊÉ}ÉXÉ^ÇçXêV
                                '-------------------------------------------------'
                                gSL_Select = "UPDATE çﬁóøéÊà¯êÊÉ}ÉXÉ^ SET "
                                If RTrim(TRM.ì`ï[éÌóﬁ) <> "2" Then
                                    gSL_Select = gSL_Select & "ì`ï[éÌóﬁ  = CASE WHEN ì`ï[éÌóﬁ = '" & RTrim(wDenpyo) & "' THEN '"
                                    gSL_Select = gSL_Select & RTrim(TRM.ì`ï[éÌóﬁ) & "' ELSE ì`ï[éÌóﬁ END ,"
                                End If
                                gSL_Select = gSL_Select & "î[ïièëå`ë‘  = CASE WHEN î[ïièëå`ë‘ = '" & RTrim(wNouKei) & "' THEN '"
                                gSL_Select = gSL_Select & RTrim(TRM.î[ïièëå`ë‘) & "' ELSE î[ïièëå`ë‘ END ,"
                                gSL_Select = gSL_Select & "éÛóÃèëå`ë‘  = CASE WHEN éÛóÃèëå`ë‘ = '" & RTrim(wJyuKei) & "' THEN '"
                                gSL_Select = gSL_Select & RTrim(TRM.éÛóÃèëå`ë‘) & "' ELSE éÛóÃèëå`ë‘ END ,"
                                gSL_Select = gSL_Select & "éÛóÃèëî≠çs  = CASE WHEN éÛóÃèëî≠çs = " & wJyuPkb & " THEN "
                                gSL_Select = gSL_Select & TRM.éÛóÃèëî≠çs & " ELSE éÛóÃèëî≠çs END ,"
                                gSL_Select = gSL_Select & "ñæç◊èëî≠çs  = CASE WHEN ñæç◊èëî≠çs = '" & RTrim(wMeiPkb) & "' THEN '"
                                gSL_Select = gSL_Select & RTrim(TRM.ñæç◊èëî≠çs) & "' ELSE ñæç◊èëî≠çs END ,"
                                gSL_Select = gSL_Select & "î[ïièëî≠çsãÊï™  = CASE WHEN î[ïièëî≠çsãÊï™ = " & wPrtkbN & " THEN "
                                gSL_Select = gSL_Select & TRM.î[ïièëî≠çsãÊï™ & " ELSE î[ïièëî≠çsãÊï™ END ,"
                                gSL_Select = gSL_Select & "éxããèëî≠çsãÊï™  = CASE WHEN éxããèëî≠çsãÊï™ = " & wPrtkbS & " THEN "
                                gSL_Select = gSL_Select & TRM.éxããèëî≠çsãÊï™ & " ELSE éxããèëî≠çsãÊï™ END ,"
                                gSL_Select = gSL_Select & "éÛóÃèëî≠çsãÊï™  = CASE WHEN éÛóÃèëî≠çsãÊï™ = " & wPrtkbJ & " THEN "
                                gSL_Select = gSL_Select & TRM.éÛóÃèëî≠çsãÊï™ & " ELSE éÛóÃèëî≠çsãÊï™ END ,"
                                gSL_Select = gSL_Select & "çXêVì˙  = GETDATE(),"
                                gSL_Select = gSL_Select & "çXêVí[ññ  = HOST_NAME(),"
                                gSL_Select = gSL_Select & "çXêVé“  = '" & gTanto & "'"
                                gSL_Select = gSL_Select & wWhere1 & wWhere3
                                With CisDB
                                    .SQL = gSL_Select
                                    On Error GoTo DBPut_Err
                                    .DBExec
                                End With
                            End If
                        End With
                    End If
                End If
            End If
        End If
    End If
'-------------------------------------------------------------------------------------------'
'   éÛì¸ÇÃÉTÉCÉNÉãÇ™ïœçXÇ≥ÇÍÇΩèÍçáÅA
'       ämîFå„ÅAïiî‘éÊà¯êÊÉ}ÉXÉ^ÇÃÉTÉCÉNÉãÇçXêV
'-------------------------------------------------------------------------------------------'
    If Not AddFlg Then
       For gCnt1 = 0 To UBound(wwUkeir)
           If Trim(wwUkeir(gCnt1)) = "" Then Exit For
            For gCnt2 = 0 To UBound(wUkeir)
                If Trim(wwUkeir(gCnt1)) = Trim(wUkeir(gCnt2)) Then
                   If Trim(wwCycle(gCnt1)) <> Trim(wCycle(gCnt2)) Then
                      wWhere1 = " WHERE éÊà¯êÊ = '" & RTrim(TRM.éÊà¯êÊCD) & "'"
                      wWhere1 = wWhere1 & " And éÛì¸ = '" & RTrim(wwUkeir(gCnt1)) & "'"
                      wWhere1 = wWhere1 & " AND ïiñ⁄ = " & RTrim(TRM.éÊà¯êÊãÊï™)
                      wWhere2 = ""
                      wWhere3 = ""
                      gStr = StrConv(wUkeir(gCnt1), vbWide) & "éÛì¸ÇÃÉTÉCÉNÉã"
                      '(( ïiî‘éÊà¯êÊÉ}ÉXÉ^çXêV ))
                      gSL_Select = "SELECT COUNT(ïiî‘) åèêî FROM ïiî‘éÊà¯êÊÉ}ÉXÉ^"
                      gSL_Select = gSL_Select & wWhere1 & wWhere2 & wWhere3
                      If HTMRead(gSL_Select, 1) Then
                          If HTM.åèêî <> 0 Then
                              With CisFun
                                  .MB_Lines = 6
                                  .MB_MSG(1) = "     " & gStr & " Ç™ïœçXÇ≥ÇÍÇ‹ÇµÇΩÅB      "
                                  .MB_MSG(3) = "     ïœçXëOÇÃílÇéùÇ¬Ç±ÇÃéÊà¯êÊÇÃïiî‘Ç™ " & HTM.åèêî & " åèë∂ç›ÇµÇ‹Ç∑ÅB      "
                                  .MB_MSG(5) = "     ïiî‘É}ÉXÉ^ÇÃÉTÉCÉNÉãÇïœçXå„ÇÃílÇ…àÍäáíuä∑ÇµÇ‹Ç∑Ç©ÅH      "
                                  .MB_Title = "ÉTÉCÉNÉãíuä∑èàóù"
                                  .MB_Button = OK_CAN
                                  If .MBOX Then
                                      '-------------------------------------------------'
                                      '   ämîFå„ÅAïiî‘éÊà¯êÊÉ}ÉXÉ^ÇçXêV
                                      '-------------------------------------------------'
                                      gSL_Select = "UPDATE ïiî‘éÊà¯êÊÉ}ÉXÉ^ SET "
                                      gSL_Select = gSL_Select & "ÉTÉCÉNÉã  = '" & wwCycle(gCnt1) & "',"
                                      gSL_Select = gSL_Select & "çXêVì˙  = GETDATE(),"
                                      gSL_Select = gSL_Select & "çXêVí[ññ  = HOST_NAME(),"
                                      gSL_Select = gSL_Select & "çXêVé“  = '" & gTanto & "'"
                                      gSL_Select = gSL_Select & wWhere1
                                      With CisDB
                                          .SQL = gSL_Select
                                          On Error GoTo DBPut_Err
                                          .DBExec
                                      End With
                                  End If
                              End With
                          End If
                      End If
                  End If
                End If
              Next gCnt2
       Next gCnt1
    '----------------------------------------------------------------------(2005.07.15)Start --
       'ójì˙ê›íËÇÃïœçX
       For gCnt1 = 1 To vsGrid.Rows - 1
            With vsGrid
                If Trim(.TextMatrix(gCnt1, 1)) = "" Then Exit For
                Erase wYoubi
                If Trim(.TextMatrix(gCnt1, 6)) = "" Then wYoubi(1) = 0 Else wYoubi(1) = 1
                If Trim(.TextMatrix(gCnt1, 7)) = "" Then wYoubi(2) = 0 Else wYoubi(2) = 1
                If Trim(.TextMatrix(gCnt1, 8)) = "" Then wYoubi(3) = 0 Else wYoubi(3) = 1
                If Trim(.TextMatrix(gCnt1, 9)) = "" Then wYoubi(4) = 0 Else wYoubi(4) = 1
                If Trim(.TextMatrix(gCnt1, 10)) = "" Then wYoubi(5) = 0 Else wYoubi(5) = 1
                
                If wYoubi(1) <> CisFun.Val2(.TextMatrix(gCnt1, 14)) Or _
                    wYoubi(2) <> CisFun.Val2(.TextMatrix(gCnt1, 15)) Or _
                    wYoubi(3) <> CisFun.Val2(.TextMatrix(gCnt1, 16)) Or _
                    wYoubi(4) <> CisFun.Val2(.TextMatrix(gCnt1, 17)) Or _
                    wYoubi(5) <> CisFun.Val2(.TextMatrix(gCnt1, 18)) Then
                    wWhere1 = " WHERE éÊà¯êÊ = '" & RTrim(TRM.éÊà¯êÊCD) & "'"
                    wWhere1 = wWhere1 & " And éÛì¸ = '" & RTrim(.TextMatrix(gCnt1, 1)) & "'"
                    wWhere1 = wWhere1 & " AND ïiñ⁄ = " & RTrim(TRM.éÊà¯êÊãÊï™)
'                    If wYoubi(1) <> CisFun.Val2(.TextMatrix(gCnt1, 14)) Then
'                        wWhere2 = "  and Week1 = " & CisFun.Val2(.TextMatrix(gCnt1, 14)) & ""
'                    End If
'                    If wYoubi(2) <> CisFun.Val2(.TextMatrix(gCnt1, 15)) Then
'                        wWhere2 = wWhere2 & "  and Week2 = " & CisFun.Val2(.TextMatrix(gCnt1, 15)) & ""
'                    End If
'                    If wYoubi(3) <> CisFun.Val2(.TextMatrix(gCnt1, 16)) Then
'                        wWhere2 = wWhere2 & "  and Week3 = " & CisFun.Val2(.TextMatrix(gCnt1, 16)) & ""
'                    End If
'                    If wYoubi(4) <> CisFun.Val2(.TextMatrix(gCnt1, 17)) Then
'                        wWhere2 = wWhere2 & "  and Week4 = " & CisFun.Val2(.TextMatrix(gCnt1, 17)) & ""
'                    End If
'                    If wYoubi(5) <> CisFun.Val2(.TextMatrix(gCnt1, 18)) Then
'                        wWhere2 = wWhere2 & "  and Week5 = " & CisFun.Val2(.TextMatrix(gCnt1, 18))
'                    End If
                    
                    wWhere3 = ""
                    gStr = StrConv(.TextMatrix(gCnt1, 1), vbWide) & "éÛì¸ÇÃójì˙î≠íçê›íË"
                    '(( ïiî‘éÊà¯êÊÉ}ÉXÉ^çXêV ))
                    gSL_Select = "SELECT COUNT(ïiî‘) åèêî FROM ïiî‘éÊà¯êÊÉ}ÉXÉ^"
                    gSL_Select = gSL_Select & wWhere1
                    If HTMRead(gSL_Select, 1) Then
                        If HTM.åèêî <> 0 Then
                            With CisFun
                                .MB_Lines = 6
                                .MB_MSG(1) = "     " & gStr & " Ç™ïœçXÇ≥ÇÍÇ‹ÇµÇΩÅB      "
                                .MB_MSG(3) = "     ïœçXëOÇÃílÇéùÇ¬Ç±ÇÃéÊà¯êÊÇÃïiî‘Ç™ " & HTM.åèêî & " åèë∂ç›ÇµÇ‹Ç∑ÅB      "
                                .MB_MSG(5) = "     ïiî‘É}ÉXÉ^ÇÃójì˙î≠íçê›íËÇïœçXå„ÇÃílÇ…àÍäáíuä∑ÇµÇ‹Ç∑Ç©ÅH      "
                                .MB_Title = "ójì˙î≠íçê›íËíuä∑èàóù"
                                .MB_Button = OK_CAN
                                If .MBOX Then
                                    '-----------------------------------------------------------'
                                    '   ämîFå„ÅAïiî‘éÊà¯êÊÉ}ÉXÉ^ÇçXêV(éÊà¯êÊÅEéÛì¸ÇÃïiî‘ëSÇƒçXêVÅj
                                    '-----------------------------------------------------------'
                                    gSL_Select = "UPDATE ïiî‘éÊà¯êÊÉ}ÉXÉ^ SET "
                                    gSL_Select = gSL_Select & "Week1  = " & wYoubi(1) & ","
                                    gSL_Select = gSL_Select & "Week2  = " & wYoubi(2) & ","
                                    gSL_Select = gSL_Select & "Week3  = " & wYoubi(3) & ","
                                    gSL_Select = gSL_Select & "Week4  = " & wYoubi(4) & ","
                                    gSL_Select = gSL_Select & "Week5  = " & wYoubi(5) & ","
                                    If Trim(vsGrid.TextMatrix(gCnt1, 5)) = "Ç∑ÇÈ" Then
                                        gSL_Select = gSL_Select & "Day1  = 0,"
                                        gSL_Select = gSL_Select & "Day2  = 0,"
                                        gSL_Select = gSL_Select & "Day3  = 0,"
                                        gSL_Select = gSL_Select & "Day4  = 0,"
                                        gSL_Select = gSL_Select & "Day5  = 0,"
                                    End If
                                    gSL_Select = gSL_Select & "çXêVì˙  = GETDATE(),"
                                    gSL_Select = gSL_Select & "çXêVí[ññ  = HOST_NAME(),"
                                    gSL_Select = gSL_Select & "çXêVé“  = '" & gTanto & "'"
                                    gSL_Select = gSL_Select & wWhere1 '& wWhere2
                                    With CisDB
                                        .SQL = gSL_Select
                                        On Error GoTo DBPut_Err
                                        .DBExec
                                    End With
                                End If
                            End With
                        End If
                    End If
                End If
            End With
       Next gCnt1
    '----------------------------------------------------------------------(2005.07.15)End --
    End If
    
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
    If Not AddFlg And wMaeHak <> TRM.ëOçHíˆãÊï™ And TRM.éÊà¯êÊãÊï™ = 1 Then
    '+------------------------------------------------------+
    '+     ëOçHíˆî≠çsãÊï™ÇÃïœçXÇ™Ç†Ç¡ÇΩèÍçá,                +
    '+     ämîFÉÅÉbÉZÅ[ÉWï\é¶å„ÅAïiî‘éÊà¯êÊÉ}ÉXÉ^ÇÃçXêV     +
    '+------------------------------------------------------+
    '=============== ìØàÍèÓïÒÇéùÇ¬ïiî‘éÊà¯êÊÉ}ÉXÉ^ÇÃämîF =================
       gSL_Select = ""
       gSL_Select = gSL_Select & " SELECT COUNT(ïiî‘) åèêî FROM ïiî‘éÊà¯êÊÉ}ÉXÉ^"
       gSL_Select = gSL_Select & "  WHERE éÊà¯êÊ = '" & RTrim(TRM.éÊà¯êÊCD) & "'"
       gSL_Select = gSL_Select & "    AND ïiñ⁄ = " & RTrim(TRM.éÊà¯êÊãÊï™)
       gSL_Select = gSL_Select & "    AND ëOçHíˆãÊï™ = " & RTrim(TRM.ëOçHíˆãÊï™)
       If HTMRead(gSL_Select, 1) Then
          If HTM.åèêî <> 0 Then
             With CisFun
                  '============== ëOçHíˆãÊï™ïœçXÇÃämîFÉÅÉbÉZÅ[ÉW ================
                  .MB_Lines = 6
                  .MB_MSG(1) = "     ëOçHíˆãÊï™Ç™ïœçXÇ≥ÇÍÇ‹ÇµÇΩÅB      "
                  .MB_MSG(3) = "     ïœçXëOÇÃílÇéùÇ¬Ç±ÇÃéÊà¯êÊÇÃïiî‘Ç™ " & HTM.åèêî & " åèë∂ç›ÇµÇ‹Ç∑ÅB      "
                  .MB_MSG(5) = "     ïiî‘É}ÉXÉ^ÇÃ" & gStr & "ÇïœçXå„ÇÃílÇ…àÍäáíuä∑ÇµÇ‹Ç∑Ç©ÅH      "
                  .MB_Title = "ëOçHíˆãÊï™íuä∑èàóù"
                  .MB_Button = OK_CAN
                  If .MBOX() Then
                     With CisDB
                         '=============== ïiî‘éÊà¯êÊÉ}ÉXÉ^çXêV =================
                         .SQL = ""
                         .SQL = .SQL & " UPDATE ïiî‘éÊà¯êÊÉ}ÉXÉ^ SET" & vbCrLf
                         .SQL = .SQL & "        ëOçHíˆãÊï™ = " & TRM.ëOçHíˆãÊï™ & "," & vbCrLf
                         .SQL = .SQL & "        çXêVì˙  = GETDATE()," & vbCrLf
                         .SQL = .SQL & "        çXêVí[ññ  = HOST_NAME()," & vbCrLf
                         .SQL = .SQL & "        çXêVé“  = '" & gTanto & "'" & vbCrLf
                         .SQL = .SQL & "  WHERE éÊà¯êÊ = '" & TRM.éÊà¯êÊ & "'" & vbCrLf
                         .SQL = .SQL & "    AND ïiñ⁄ = " & TRM.éÊà¯êÊãÊï™ & vbCrLf
                         .SQL = .SQL & "    AND ëOçHíˆãÊï™ = " & TRM.ëOçHíˆãÊï™ & vbCrLf
                       
                         On Error GoTo DBPut_Err
                         .DBExec
                     End With
                  End If
             End With
          End If
       End If
       
    '+------------------------------------------------------+
    '+     ëOçHíˆî≠çsãÊï™ÇÃïœçXÇ™Ç†Ç¡ÇΩèÍçá,                +
    '+     ämîFÉÅÉbÉZÅ[ÉWï\é¶å„ÅAçﬁóøéÊà¯êÊÉ}ÉXÉ^ÇÃçXêV     +
    '+------------------------------------------------------+
    '=============== ìØàÍèÓïÒÇéùÇ¬çﬁóøéÊà¯êÊÉ}ÉXÉ^ÇÃämîF =================
       gSL_Select = ""
       gSL_Select = gSL_Select & " SELECT COUNT(çﬁóøä«óùî‘çÜ) åèêî FROM çﬁóøéÊà¯êÊÉ}ÉXÉ^"
       gSL_Select = gSL_Select & "  WHERE éÊà¯êÊ = '" & RTrim(TRM.éÊà¯êÊCD) & "'"
       gSL_Select = gSL_Select & "    AND ïiñ⁄ = " & RTrim(TRM.éÊà¯êÊãÊï™)
       gSL_Select = gSL_Select & "    AND ëOçHíˆãÊï™ = " & RTrim(TRM.ëOçHíˆãÊï™)
       If HTMRead(gSL_Select, 1) Then
          If HTM.åèêî <> 0 Then
             With CisFun
                  '============== ëOçHíˆãÊï™ïœçXÇÃämîFÉÅÉbÉZÅ[ÉW ================
                  .MB_Lines = 6
                  .MB_MSG(1) = "     ëOçHíˆãÊï™Ç™ïœçXÇ≥ÇÍÇ‹ÇµÇΩÅB      "
                  .MB_MSG(3) = "     ïœçXëOÇÃílÇéùÇ¬Ç±ÇÃéÊà¯êÊÇÃçﬁóøÇ™ " & HTM.åèêî & " åèë∂ç›ÇµÇ‹Ç∑ÅB      "
                  .MB_MSG(5) = "     çﬁóøÉ}ÉXÉ^ÇÃ" & gStr & "ÇïœçXå„ÇÃílÇ…àÍäáíuä∑ÇµÇ‹Ç∑Ç©ÅH      "
                  .MB_Title = "ëOçHíˆãÊï™íuä∑èàóù"
                  .MB_Button = OK_CAN
                  If .MBOX() Then
                     With CisDB
                          '=============== ïiî‘éÊà¯êÊÉ}ÉXÉ^çXêV =================
                          .SQL = ""
                          .SQL = .SQL & " UPDATE çﬁóøéÊà¯êÊÉ}ÉXÉ^ SET" & vbCrLf
                          .SQL = .SQL & "        ëOçHíˆãÊï™ = " & TRM.ëOçHíˆãÊï™ & "," & vbCrLf
                          .SQL = .SQL & "        çXêVì˙  = GETDATE()," & vbCrLf
                          .SQL = .SQL & "        çXêVí[ññ  = HOST_NAME()," & vbCrLf
                          .SQL = .SQL & "        çXêVé“  = '" & gTanto & "'" & vbCrLf
                          .SQL = .SQL & "  WHERE éÊà¯êÊ = '" & TRM.éÊà¯êÊ & "'" & vbCrLf
                          .SQL = .SQL & "    AND ïiñ⁄ = " & TRM.éÊà¯êÊãÊï™ & vbCrLf
                          .SQL = .SQL & "    AND ëOçHíˆãÊï™ = " & TRM.ëOçHíˆãÊï™ & vbCrLf
                           
                          On Error GoTo DBPut_Err
                          .DBExec
                      End With
                  End If
             End With
          End If
       End If
       
    '================================================================== *** 2009/07/06 Start
       '------------------------------
       '-- èoâ◊èÍàÍäáïœçX
       '--   Å¶éÛíç ÇÃÇ›ëŒè€
       '------------------------------
       If H1Op_Torkb(0).Value Then
          If RTrim(TRM.èoâ◊èÍ) <> RTrim(wSyukaba) Then
              wWhere1 = " WHERE éÊà¯êÊ = '" & RTrim(TRM.éÊà¯êÊCD) & "'"
              wWhere1 = wWhere1 & " AND ïiñ⁄ = 0"
              wWhere1 = wWhere1 & " AND èoâ◊èÍ = '" & RTrim(wSyukaba) & "'"
              '(( ïiî‘éÊà¯êÊÉ}ÉXÉ^çXêV ))
              gSL_Select = "SELECT COUNT(ïiî‘) åèêî FROM ïiî‘éÊà¯êÊÉ}ÉXÉ^"
              gSL_Select = gSL_Select & wWhere1
          
              If HTMRead(gSL_Select, 1) Then
                  If HTM.åèêî <> 0 Then
                      With CisFun
                          .MB_Lines = 6
                          .MB_MSG(1) = "     Åyèoâ◊èÍÅz Ç™ïœçXÇ≥ÇÍÇ‹ÇµÇΩÅB      "
                          .MB_MSG(3) = "     ïœçXëOÇÃílÇéùÇ¬Ç±ÇÃéÊà¯êÊÇÃïiî‘Ç™ " & HTM.åèêî & " åèë∂ç›ÇµÇ‹Ç∑ÅB      "
                          .MB_MSG(5) = "     ïiî‘É}ÉXÉ^ÇÃÅyèoâ◊èÍÅzÇïœçXå„ÇÃílÇ…àÍäáíuä∑ÇµÇ‹Ç∑Ç©ÅH      "
                          .MB_Title = "Åyèoâ◊èÍÅzíuä∑èàóù"
                          .MB_Button = OK_CAN
                          If .MBOX Then
                              '-------------------------------------------------'
                              '   ämîFå„ÅAïiî‘éÊà¯êÊÉ}ÉXÉ^ÇçXêV
                              '-------------------------------------------------'
                              gSL_Select = "UPDATE ïiî‘éÊà¯êÊÉ}ÉXÉ^ SET "
                              gSL_Select = gSL_Select & "èoâ◊èÍ  = '" & RTrim(iB1_Syukaba) & "',"
                              gSL_Select = gSL_Select & "çXêVì˙  = GETDATE(),"
                              gSL_Select = gSL_Select & "çXêVí[ññ  = HOST_NAME(),"
                              gSL_Select = gSL_Select & "çXêVé“  = '" & gTanto & "'"
                              gSL_Select = gSL_Select & wWhere1
                              With CisDB
                                  .SQL = gSL_Select
                                  On Error GoTo DBPut_Err
                                  .DBExec
                              End With
                          End If
                      End With
                  End If
              End If
          End If
       End If
    '================================================================== *** 2009/07/06 End
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
    
    CisDB.DBTran (TransCommit) ' ƒ◊›ªﬁ∏ºÆ› èIóπ
    
    DBPut = True
    Exit Function
DBPut_Err:
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
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 2, 0, "NO")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 2, 1, "éÛ")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 21, 2, "éÛì¸ñºèÃ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 8, 3, "ÉTÉCÉNÉã")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 6, 4, "ãÛñæç◊")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 8, 5, "ójì˙î≠íç")
    '----------------------------------------------------------------------(2005.07.15)Start --
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 4, 6, "åé")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 4, 7, "âŒ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 4, 8, "êÖ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 4, 9, "ñÿ")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 4, 10, "ã‡")
    '----------------------------------------------------------------------(2005.07.15)End --
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 11, "CYCLE1")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 12, "CYCLE2")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 13, "CYCLE3")
    '----------------------------------------------------------------------(2005.07.15)Start --
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 14, "ójì˙1")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 15, "ójì˙2")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 16, "ójì˙3")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 17, "ójì˙4")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 18, "ójì˙5")
    '----------------------------------------------------------------------(2005.07.15)End --
'Å@∏ﬁÿØƒﬁëIëèàóùópÅ@ëﬁî¥ÿ±
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 0, 19, "Select")
'  ï\é¶ ﬂ¿∞›æØƒ
'       Call .PatanSet( ﬂ¿∞›NO,ï\é¶ÉpÉ^Å[Éì)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10")
        .InitDisp
    End With
End Sub
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
    
'ÉeÉLÉXÉgÇ©ÇÁÉOÉäÉbÉhÇ…ëﬁî
    Call TextToVsGrid(Val(iB1_Ukeir.Tag))
    gInt = vsGrid.Row
    'Å™
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    'Å´
    If KeyCode = vbKeyDown Then
        If gInt + 1 > vsGrid.Rows - 1 Then
            If Trim(vsGrid.TextMatrix(vsGrid.Rows - 1, 1)) = "" Then GoTo Grid_RowChange_End
            vsGrid.Rows = vsGrid.Rows + 1
        End If
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
'    If KeyCode = vbKeyPageDown Then
'        If gInt + CisVsGrid.DispMax > Ttl_Kensu - 1 Then
'            gInt = Ttl_Kensu - 1
'        Else
'            gInt = gInt + CisVsGrid.DispMax
'        End If
'    End If
    vsGrid.Row = gInt
'ÉOÉäÉbÉhÇ©ÇÁÉeÉLÉXÉgÇ…ï\é¶
    Call VsGridToText(vsGrid.Row)
    
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   ì¸óÕÉeÉLÉXÉgóLå¯ÅEñ≥å¯
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
    iB1_Ukeir.Visible = EnabledFlg
    iB1_UName.Visible = EnabledFlg
    
    If B1Chk_Torkb(1).Value = 1 Then
       iB1_Cycle1.Visible = False
       iB1_Cycle2.Visible = False
       iB1_Cycle3.Visible = False
       ChkB1_Kara.Visible = False
       ChkB1_YoubiHC.Visible = False
       ChkB1_Youbi(0).Visible = False
       ChkB1_Youbi(1).Visible = False
       ChkB1_Youbi(2).Visible = False
       ChkB1_Youbi(3).Visible = False
       ChkB1_Youbi(4).Visible = False
       iB1_Cycle1 = ""
       iB1_Cycle2 = ""
       iB1_Cycle3 = ""
       ChkB1_Kara.Caption = ""
    Else
        iB1_Cycle1.Visible = EnabledFlg
        iB1_Cycle2.Visible = EnabledFlg
        iB1_Cycle3.Visible = EnabledFlg
        If H1Op_Torkb(1).Value Then
            ChkB1_Kara.Visible = EnabledFlg
            If Trim(ChkB1_Kara.Caption) = "" Then
               ChkB1_Kara.Caption = "óL"
            End If
    '---------------------------------------------------------------- 2006/02/23 Insert
            If Trim(ChkB1_YoubiHC.Caption) = "" Then
               ChkB1_YoubiHC.Caption = "ÇµÇ»Ç¢"
            End If
    '---------------------------------------------------------------- 2006/02/23 Insert End

            ChkB1_YoubiHC.Visible = EnabledFlg
            ChkB1_Youbi(0).Visible = EnabledFlg
            ChkB1_Youbi(1).Visible = EnabledFlg
            ChkB1_Youbi(2).Visible = EnabledFlg
            ChkB1_Youbi(3).Visible = EnabledFlg
            ChkB1_Youbi(4).Visible = EnabledFlg
        Else
            ChkB1_Kara.Caption = ""
            ChkB1_Kara.Visible = False
            ChkB1_YoubiHC.Caption = ""
            ChkB1_YoubiHC.Visible = False
            ChkB1_Youbi(0).Visible = False
            ChkB1_Youbi(1).Visible = False
            ChkB1_Youbi(2).Visible = False
            ChkB1_Youbi(3).Visible = False
            ChkB1_Youbi(4).Visible = False
        End If
    End If
End Sub
'==========================================================='
'   ì¸óÕÉeÉLÉXÉgÇ©ÇÁÉOÉäÉbÉhÇ…à⁄Ç∑
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    With vsGrid
        .TextMatrix(tRow, 1) = iB1_Ukeir
        If Trim(iB1_Ukeir) = "" Then
            For gInt = 2 To .Cols - 1
                .TextMatrix(tRow, gInt) = ""
            Next gInt
            Exit Sub
        End If
        .TextMatrix(tRow, 0) = tRow
        .TextMatrix(tRow, 2) = iB1_UName
        .TextMatrix(tRow, 11) = iB1_Cycle1
        .TextMatrix(tRow, 12) = iB1_Cycle2
        .TextMatrix(tRow, 13) = iB1_Cycle3
        .TextMatrix(tRow, 3) = CycleFormat(tRow)
        .TextMatrix(tRow, 4) = ChkB1_Kara.Caption
    '----------------------------------------------------------------------(2005.07.15)Start --
        .TextMatrix(tRow, 5) = ChkB1_YoubiHC.Caption
        .TextMatrix(tRow, 6) = ChkB1_Youbi(0).Caption
        .TextMatrix(tRow, 7) = ChkB1_Youbi(1).Caption
        .TextMatrix(tRow, 8) = ChkB1_Youbi(2).Caption
        .TextMatrix(tRow, 9) = ChkB1_Youbi(3).Caption
        .TextMatrix(tRow, 10) = ChkB1_Youbi(4).Caption
    '----------------------------------------------------------------------(2005.07.15)End --
    End With
End Sub
'==========================================================='
'   ÉOÉäÉbÉhÇ©ÇÁì¸óÕÉeÉLÉXÉgÇ…à⁄Ç∑ÅAì¸óÕÉeÉLÉXÉgÇÃà⁄ìÆ
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With vsGrid
        'éÛì¸
        iB1_Ukeir.Tag = tRow
        iB1_Ukeir = RTrim(.TextMatrix(tRow, 1))
        iB1_Ukeir.Move iB1_Ukeir.Left, vsGrid.CellTop, iB1_Ukeir.Width, vsGrid.CellHeight
        'éÛì¸ñº
        iB1_UName = RTrim(.TextMatrix(tRow, 2))
        iB1_UName.Move iB1_UName.Left, vsGrid.CellTop, iB1_UName.Width, vsGrid.CellHeight
        ' ÉTÉCÉNÉã
        iB1_Cycle1 = CisFun.RSetFld(.TextMatrix(tRow, 11), 0, iB1_Cycle1.cFormat)
        iB1_Cycle1.Move iB1_Cycle1.Left, vsGrid.CellTop, iB1_Cycle1.Width, vsGrid.CellHeight
        iB1_Cycle2 = CisFun.RSetFld(.TextMatrix(tRow, 12), 0, iB1_Cycle2.cFormat)
        iB1_Cycle2.Move iB1_Cycle2.Left, vsGrid.CellTop, iB1_Cycle2.Width, vsGrid.CellHeight
        iB1_Cycle3 = CisFun.RSetFld(.TextMatrix(tRow, 13), 0, iB1_Cycle3.cFormat)
        iB1_Cycle3.Move iB1_Cycle3.Left, vsGrid.CellTop, iB1_Cycle3.Width, vsGrid.CellHeight
        'ãÛñæç◊
        If RTrim(.TextMatrix(tRow, 4)) = "ñ≥" Then
           ChkB1_Kara.Value = 1
        Else
           ChkB1_Kara.Value = 0
        End If
        ChkB1_Kara.Move ChkB1_Kara.Left, vsGrid.CellTop, ChkB1_Kara.Width, vsGrid.CellHeight
    '----------------------------------------------------------------------(2005.07.15)Start --
        'ójì˙î≠íç
        If RTrim(.TextMatrix(tRow, 5)) = "Ç∑ÇÈ" Then
           ChkB1_YoubiHC.Value = 1
        Else
           ChkB1_YoubiHC.Value = 0
        End If
        ChkB1_YoubiHC.Move ChkB1_YoubiHC.Left, vsGrid.CellTop, ChkB1_YoubiHC.Width, vsGrid.CellHeight
        'ójì˙ê›íË
        For gInt = 0 To 4
            If RTrim(.TextMatrix(tRow, 6 + gInt)) = "Åõ" Then
               ChkB1_Youbi(gInt).Value = 1
            Else
               ChkB1_Youbi(gInt).Value = 0
            End If
            ChkB1_Youbi(gInt).Move ChkB1_Youbi(gInt).Left, vsGrid.CellTop, ChkB1_Youbi(gInt).Width, vsGrid.CellHeight
        Next gInt
    '----------------------------------------------------------------------(2005.07.15)End --
    End With
End Sub
'******************************************'
'*****       ÉOÉäÉbÉhì‡óeÉZÉbÉg        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Long
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
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not TUM_RDSTS
        With vsGrid
            ID = ID + 1
            'áÇ
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            'éÛì¸
            .TextMatrix(ID, GridCol(1)) = TUM.éÛì¸
            ' éÛì¸ñºèÃ
            .TextMatrix(ID, GridCol(2)) = TUM.éÛì¸ñºèÃ
            If Trim(TUM.ÉTÉCÉNÉã) <> "" Then
                .TextMatrix(ID, GridCol(11)) = Mid(TUM.ÉTÉCÉNÉã, 1, 1)
                .TextMatrix(ID, GridCol(12)) = Mid(TUM.ÉTÉCÉNÉã, 2, 2)
                .TextMatrix(ID, GridCol(13)) = Mid(TUM.ÉTÉCÉNÉã, 4, 2)
                .TextMatrix(ID, GridCol(3)) = CycleFormat(ID)
            End If
            If TUM.éÊà¯êÊãÊï™ <> 0 Then
                ' ãÛñæç◊
                If TUM.ãÛñæç◊ = "1" Then
                   .TextMatrix(ID, GridCol(4)) = "ñ≥"
                Else
                   .TextMatrix(ID, GridCol(4)) = "óL"
                End If
        '----------------------------------------------------------------------(2005.07.15)Start --
                ' ójì˙î≠íç
                If TUM.ójì˙î≠íç = 0 Then
                   .TextMatrix(ID, GridCol(5)) = "ÇµÇ»Ç¢"
                Else
                   .TextMatrix(ID, GridCol(5)) = "Ç∑ÇÈ"
                End If
                ' ójì˙ê›íË
                If TUM.ójì˙1 <> 0 Then
                   .TextMatrix(ID, GridCol(6)) = "Åõ"
                End If
                If TUM.ójì˙2 <> 0 Then
                   .TextMatrix(ID, GridCol(7)) = "Åõ"
                End If
                If TUM.ójì˙3 <> 0 Then
                   .TextMatrix(ID, GridCol(8)) = "Åõ"
                End If
                If TUM.ójì˙4 <> 0 Then
                   .TextMatrix(ID, GridCol(9)) = "Åõ"
                End If
                If TUM.ójì˙5 <> 0 Then
                   .TextMatrix(ID, GridCol(10)) = "Åõ"
                End If
                .TextMatrix(ID, GridCol(14)) = Format(TUM.ójì˙1, "0")
                .TextMatrix(ID, GridCol(15)) = Format(TUM.ójì˙2, "0")
                .TextMatrix(ID, GridCol(16)) = Format(TUM.ójì˙3, "0")
                .TextMatrix(ID, GridCol(17)) = Format(TUM.ójì˙4, "0")
                .TextMatrix(ID, GridCol(18)) = Format(TUM.ójì˙5, "0")
            End If
    '----------------------------------------------------------------------(2005.07.15)End --
        End With
        Call TUMReadNext(1)
    Loop
    
    With vsGrid
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++     √∑Ωƒì¸óÕéû      +++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub vsGrid_Scroll()
    'ÉXÉNÉçÅ[ÉãéûÇÕÉeÉLÉXÉgÉ{ÉbÉNÉXÇîÒï\é¶Ç…ÇµÇ‹Ç∑ÅB
'    Call GridTextEnabled(False)
End Sub
Private Sub VSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown Then Exit Sub
    Call TextToVsGrid(Val(iB1_Ukeir.Tag))
    Call VsGridToText(vsGrid.Row)
End Sub
Private Sub VSGrid_Click()
'    Call GridTextEnabled(True)
    Call TextToVsGrid(Val(iB1_Ukeir.Tag))
    Call VsGridToText(vsGrid.Row)
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub
'+---------------------------------------+
'+      ñºèÃÉ}ÉXÉ^èâä˙ê›íË
'+---------------------------------------+
Private Sub MeisyoInit()
' ì`ï[éÌóﬁî≠íç
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "ì`ï[éÌóﬁî≠íç"
        .ãÊï™åÖêî = 1
        .íl = "1"
        .êßå‰ãÊï™ = "X"
        .ílñºèÃ = "é©ããJAMA"
        .çÏê¨é“ = gTanto
        .åüçıñº = "î≠íçì`ï[éÌóﬁ åüçı"
        .ÉRÅ[Éhñº = "ì`ï[éÌ"
        .ñºèÃñº = "ì`ï[éÌóﬁñº"
    End With
    If Not SYMInitCreate(True) Then
        With SYMCreate
            .íl = "2"
            .ílñºèÃ = "éxããJAMA"
            .êîéö1 = 1
        End With
        Call SYMInitCreate
    '----<< É`ÉPÉbÉgí«â¡  >>-------------------------------------------------------------- 2005.01.31 (START)-------'
        With SYMCreate
            .íl = "3"
            .ílñºèÃ = "é©ãã¡πØƒ"
            .êîéö1 = 0
        End With
        Call SYMInitCreate
        With SYMCreate
            .íl = "4"
            .ílñºèÃ = "éxãã¡πØƒ"
            .êîéö1 = 1
        End With
        Call SYMInitCreate
    '----<< É`ÉPÉbÉgí«â¡  >>-------------------------------------------------------------- 2005.01.31 (END  )-------'
    End If
' ì`ï[éÌóﬁéÛíç
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "ì`ï[éÌóﬁéÛíç"
        .ãÊï™åÖêî = 1
        .íl = "1"
        .êßå‰ãÊï™ = "X"
        .ílñºèÃ = "é©é–ì`ï["
        .çÏê¨é“ = gTanto
        .åüçıñº = "éÛíçì`ï[éÌóﬁ åüçı"
        .ÉRÅ[Éhñº = "ì`ï[éÌ"
        .ñºèÃñº = "ì`ï[éÌóﬁñº"
    End With
    If Not SYMInitCreate(True) Then
        With SYMCreate
            .íl = "2"
            .ílñºèÃ = "ÇiÇ`ÇlÇ`"
        End With
        Call SYMInitCreate
    End If
' î[ïièëå`ë‘î≠íç
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "î[ïièëå`ë‘î≠íç"
        .ãÊï™åÖêî = 1
        .íl = "1"
        .êßå‰ãÊï™ = "X"
        .ílñºèÃ = "î≠çs"
        .çÏê¨é“ = gTanto
    End With
    If Not SYMInitCreate(True) Then
        With SYMCreate
            .íl = "2"
            .ílñºèÃ = "Ãß≤ŸèoóÕ"
        End With
        Call SYMInitCreate
''        With SYMCreate
''            .íl = "3"
''            .ílñºèÃ = "î≠çs&ëóêM"
''        End With
''        Call SYMInitCreate
    End If
' î[ïièëå`ë‘éÛíç
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "î[ïièëå`ë‘éÛíç"
        .ãÊï™åÖêî = 1
        .íl = "1"
        .êßå‰ãÊï™ = "X"
        .ílñºèÃ = "î≠çs"
        .çÏê¨é“ = gTanto
    End With
    If Not SYMInitCreate(True) Then
        With SYMCreate
            .íl = "2"
            .ílñºèÃ = "Ãß≤ŸèoóÕ"
        End With
        Call SYMInitCreate
''        With SYMCreate
''            .íl = "3"
''            .ílñºèÃ = "î≠çs&ëóêM"
''        End With
''        Call SYMInitCreate
    End If
' éÛóÃèëå`ë‘
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "éÛóÃèëå`ë‘"
        .ãÊï™åÖêî = 1
        .íl = "1"
        .êßå‰ãÊï™ = "X"
        .ílñºèÃ = "î≠çs"
        .çÏê¨é“ = gTanto
    End With
    If Not SYMInitCreate(True) Then
        With SYMCreate
            .íl = "2"
            .ílñºèÃ = "Ãß≤ŸèoóÕ"
        End With
        Call SYMInitCreate
''        With SYMCreate
''            .íl = "3"
''            .ílñºèÃ = "î≠çs&ëóêM"
''        End With
''        Call SYMInitCreate
    End If
' éÛóÃèëå`ë‘(éÛíç)
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "éÛóÃèëå`ë‘éÛíç"
        .ãÊï™åÖêî = 1
        .íl = "1"
        .êßå‰ãÊï™ = "X"
        .ílñºèÃ = "î≠çs"
        .çÏê¨é“ = gTanto
    End With
    If Not SYMInitCreate(True) Then
        With SYMCreate
            .íl = "2"
            .ílñºèÃ = "Ãß≤ŸèoóÕ"
        End With
        Call SYMInitCreate
''        With SYMCreate
''            .íl = "3"
''            .ílñºèÃ = "î≠çs&ëóêM"
''        End With
''        Call SYMInitCreate
    End If
' î≠íçì‡é¶
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "î≠íçì‡é¶"
        .ãÊï™åÖêî = 1
        .íl = "1"
        .êßå‰ãÊï™ = "X"
        .ílñºèÃ = "î≠çsÇ»Çµ"
        .çÏê¨é“ = gTanto
    End With
    If Not SYMInitCreate(True) Then
        With SYMCreate
            .íl = "2"
            .ílñºèÃ = "î≠çsÇ†ÇË"
        End With
        Call SYMInitCreate
        
        With SYMCreate
            .íl = "4"
            .ílñºèÃ = "ëóêM(Mail)"
        End With
        Call SYMInitCreate
    End If
' ñæç◊èëî≠çs
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "ñæç◊èëî≠çs"
        .ãÊï™åÖêî = 1
        .íl = "1"
        .êßå‰ãÊï™ = "X"
        .ílñºèÃ = "î≠çsÇ†ÇË"
        .çÏê¨é“ = gTanto
    End With
    If Not SYMInitCreate(True) Then
        For gInt = 2 To 4
            With SYMCreate
                .íl = Format(gInt, "0")
                Select Case gInt
                     Case 2: .ílñºèÃ = "î≠çsÇ»Çµ"
                     Case 3: .ílñºèÃ = "ëóêM(FAX)"
                     Case 4: .ílñºèÃ = "ëóêM(Mail)"
                End Select
            End With
            Call SYMInitCreate
        Next gInt
    End If
' óaã‡ãÊï™
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "óaã‡ãÊï™"
        .ãÊï™åÖêî = 1
        .íl = "1"
        .ílñºèÃ = "ïÅí "
        .çÏê¨é“ = gTanto
        .åüçıñº = "óa ã‡ ãÊ ï™ åü çı"
        .ÉRÅ[Éhñº = "óaã‡ãÊ"
        .ñºèÃñº = "óaã‡ãÊï™ñº"
    End With
    If Not SYMInitCreate(True) Then
        With SYMCreate
            .íl = "2"
            .ílñºèÃ = "ìñç¿"
        End With
        Call SYMInitCreate
    End If


'===================================== 2009/07/15 Start
' ã@äB
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "ã@äB"
        .ãÊï™åÖêî = 3
        .íl = "ZZZ"
        .ílñºèÃ = "ÇªÇÃëº"
        .çÏê¨é“ = gTanto
        .åüçıñº = "ã@ äB åü çı"
        .ÉRÅ[Éhñº = "ã@äB"
        .ñºèÃñº = "ã@äB"
    End With
    If Not SYMInitCreate(True) Then
    End If

' ã@äBNO
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "ã@äBNO"
        .ãÊï™åÖêî = 3
        .íl = "ZZZ"
        .ílñºèÃ = "ÇªÇÃëº"
        .çÏê¨é“ = gTanto
        .åüçıñº = "ã@ äB áÇ åü çı"
        .ÉRÅ[Éhñº = "ã@äBáÇ"
        .ñºèÃñº = "ã@äBáÇ"
    End With
    If Not SYMInitCreate(True) Then
    End If

' ïîñÂ1
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "ïîñÂ1"
        .ãÊï™åÖêî = 4
        .íl = "ZZZZ"
        .ílñºèÃ = "ÇªÇÃëº"
        .çÏê¨é“ = gTanto
        .åüçıñº = "ïî ñÂ 1 åü çı"
        .ÉRÅ[Éhñº = "ïîñÂ1"
        .ñºèÃñº = "ïîñÂ1"
    End With
    If Not SYMInitCreate(True) Then
    End If

' ïîñÂ2
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "ïîñÂ2"
        .ãÊï™åÖêî = 4
        .íl = "ZZZZ"
        .ílñºèÃ = "ÇªÇÃëº"
        .çÏê¨é“ = gTanto
        .åüçıñº = "ïî ñÂ 2 åü çı"
        .ÉRÅ[Éhñº = "ïîñÂ2"
        .ñºèÃñº = "ïîñÂ2"
    End With
    If Not SYMInitCreate(True) Then
    End If
'===================================== 2009/07/15 End

End Sub


'===================================================================================== 2009/01/06 Start
'****************************
'*      Grid1¡™Ø∏èàóùÅiïœä∑éÊà¯êÊ∫∞ƒﬁÅj
'****************************
Private Function Grid1Chk() As Boolean
    Dim wUke_Flg   As Boolean

    Grid1Chk = False
    
    With vsGrid1
        For gCnt1 = 1 To .Rows - 1
            If Trim(.TextMatrix(gCnt1, 1)) <> "" Or Trim(.TextMatrix(gCnt1, 2)) <> "" Or Trim(.TextMatrix(gCnt1, 3)) <> "" Then
                'ïœä∑å≥CD/ïœä∑å≥éÛì¸ÇÃèdï°É`ÉFÉbÉN
                For gCnt2 = gCnt1 + 1 To .Rows - 1
                    If Trim(.TextMatrix(gCnt2, 1)) & Trim(.TextMatrix(gCnt2, 2)) <> "" Then
                        If Trim(.TextMatrix(gCnt1, 1)) & Trim(.TextMatrix(gCnt1, 2)) = Trim(.TextMatrix(gCnt2, 1)) & Trim(.TextMatrix(gCnt2, 2)) Then
                            Set gObj = iB1_MotoCD
                            gLong = gCnt2
                            With CisFun
                                .MB_Lines = 4
                                .MB_MSG(2) = "    " & StrConv(gCnt2, vbWide) & "çsñ⁄ ì¸óÕì‡óeÇ™èdï°ÇµÇƒÇ‹Ç∑ÅB            "
                                .MB_Title = ""
                                .MB_Button = Error
                                .MBOX
                            End With
                            GoTo Grid1Chk_Err
                        End If
                    End If
                Next gCnt2
                ' ïœä∑å≥CDì¸óÕ¡™Ø∏
                If Trim(.TextMatrix(gCnt1, 1)) = "" And (Trim(.TextMatrix(gCnt1, 2)) <> "" Or Trim(.TextMatrix(gCnt1, 4)) <> "") Then
                   gLong = gCnt1
                   Set gObj = iB1_MotoCD
                   With CisFun
                        .MB_Lines = 4
                        .MB_MSG(2) = "  " & StrConv(gCnt1, vbWide) & "çsñ⁄  ïœä∑å≥CDÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
                        .MB_Title = ""
                        .MB_Button = Error
                        .MBOX
                    End With
                    GoTo Grid1Chk_Err
                End If
                ' ïœä∑å≥éÛì¸ì¸óÕ¡™Ø∏
                If Trim(.TextMatrix(gCnt1, 2)) = "" And Trim(.TextMatrix(gCnt1, 4)) <> "" Then
                   gLong = gCnt1
                   Set gObj = iB1_MotoUke
                   With CisFun
                        .MB_Lines = 4
                        .MB_MSG(2) = "  " & StrConv(gCnt1, vbWide) & "çsñ⁄  ïœä∑å≥éÛì¸Çì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
                        .MB_Title = ""
                        .MB_Button = Error
                        .MBOX
                    End With
                    GoTo Grid1Chk_Err
                End If
                ' ïœä∑êÊéÛì¸ì¸óÕ¡™Ø∏
                If Trim(.TextMatrix(gCnt1, 2)) <> "" And Trim(.TextMatrix(gCnt1, 4)) = "" Then
                   gLong = gCnt1
                   Set gObj = iB1_ChgUke
                   With CisFun
                        .MB_Lines = 4
                        .MB_MSG(2) = "  " & StrConv(gCnt1, vbWide) & "çsñ⁄  ïœä∑éÛì¸Çì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
                        .MB_Title = ""
                        .MB_Button = Error
                        .MBOX
                    End With
                    GoTo Grid1Chk_Err
                End If
                ' å_ñÒêÊì¸óÕ¡™Ø∏
'                If Trim(.TextMatrix(gCnt1, 5)) = "" And (Trim(.TextMatrix(gCnt1, 1)) <> "" Or Trim(.TextMatrix(gCnt1, 4)) <> "") Then
'                   gLong = gCnt1
'                   Set gObj = iB1_HenKeiyaku
'                   With CisFun
'                        .MB_Lines = 4
'                        .MB_MSG(2) = "  " & StrConv(gCnt1, vbWide) & "çsñ⁄  å_ñÒêÊÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
'                        .MB_Title = ""
'                        .MB_Button = Error
'                        .MBOX
'                    End With
'                    GoTo Grid1Chk_Err
'                End If
                'å_ñÒêÊìoò^çœÉ`ÉFÉbÉN
                If Trim(.TextMatrix(gCnt1, 5)) <> "" Then
                    If TorNmGet(Trim(.TextMatrix(gCnt1, 5)), 2) Then
                    Else
                        gLong = gCnt1
                        Set gObj = iB1_HenKeiyaku
                        With CisFun
                            .MB_Lines = 4
                            .MB_MSG(2) = "    [ïœä∑]å_ñÒêÊñ¢ìoò^            "
                            .MB_Title = ""
                            .MB_Button = Error
                            .MBOX
                        End With
                       Exit Function
                       GoTo Grid1Chk_Err
                    End If
                End If
            
                ' ïœä∑êÊéÛì¸ë∂ç›¡™Ø∏
                wUke_Flg = False
                For gCnt2 = 1 To vsGrid.Rows - 1
                    If Trim(.TextMatrix(gCnt1, 4)) = Trim(vsGrid.TextMatrix(gCnt2, 1)) Then
                       wUke_Flg = True
                       Exit For
                    End If
                Next gCnt2
                If Not wUke_Flg Then
                   gLong = gCnt1
                   Set gObj = iB1_ChgUke
                   With CisFun
                        .MB_Lines = 5
                        .MB_MSG(1) = "    " & StrConv(gCnt1, vbWide) & "çsñ⁄ ïœä∑êÊéÛì¸Ç™éÛì¸èÓïÒñ¢ìoò^Ç≈Ç∑ÅB      "
                        .MB_MSG(3) = "     èàóùÇë±çsÇµÇ‹Ç∑Ç©ÅH      "
                        .MB_Title = "ämîF"
                        .MB_Button = OK_CAN
                        If Not .MBOX() Then
                           GoTo Grid1Chk_Err
                        End If
                   End With
                End If
                
            End If
        Next gCnt1
    End With
    
    Grid1Chk = True
    Exit Function

Grid1Chk_Err:
    vsGrid1.Row = gLong
    Call TextToVsGrid1(Val(iB1_MotoCD.Tag))
    Call VsGridToText1(vsGrid1.Row)
    gObj.SetFocus

End Function
Private Sub VsGrid1_DBInsert()
    
    With CisDB
        .SQL = "DELETE éÛíçéÊà¯êÊCDïœä∑É}ÉXÉ^ "
        .SQL = .SQL & " WHERE ïœä∑êÊCD = '" & RTrim(iH1_Torcd) & RTrim(iH1_Kojyo) & "'"
        .DBExec
    End With

    With vsGrid1
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, 1)) & Trim(.TextMatrix(gLong, 2)) <> "" Then
               Call ItemsClearJTH
               With JTH
                    .ïœä∑å≥CD = Trim(vsGrid1.TextMatrix(gLong, 1))
                    .ïœä∑å≥éÛì¸ = Trim(vsGrid1.TextMatrix(gLong, 2))
                    .ïœä∑êÊCD = RTrim(iH1_Torcd) & RTrim(iH1_Kojyo)
                    .ïœä∑êÊéÛì¸ = Trim(vsGrid1.TextMatrix(gLong, 4))
                    .å_ñÒêÊ = Trim(vsGrid1.TextMatrix(gLong, 5))
               End With
               Call JTHInsert
            End If
        Next gLong
    End With

End Sub


'+==========================================================+
'+      éÊà¯êÊïœä∑∫∞ƒﬁì¸óÕÉOÉäÉbÉhèàóù                      +
'+==========================================================+
'+----------------------------------------------------------+
'+   å©èoÇµæØƒ & èâä˙èÛë‘æØƒ
'+----------------------------------------------------------+
Private Sub GridInit1()
    With CisVsGrid1
'Å@∏ﬁÿØƒﬁèâä˙âªèàóù
        .Clear
'       Call .FixedSet(å©èoÇµà íuê›íË,óÒà íuê›íË,œ∞ºﬁèàóù,
'                      æŸåÖêî,óÒî‘çÜ,å©èoÇµì‡óe,ï™óﬁ¿≤Ãﬂ)
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 2, 0, "NO")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 8, 1, "ïœä∑å≥CD")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 4, 2, "éÛì¸")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 2, 3, "Å@")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 8, 4, "ïœä∑éÛì¸")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 8, 5, "å_ñÒêÊ")
'Å@∏ﬁÿØƒﬁëIëèàóùópÅ@ëﬁî¥ÿ±
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 0, 6, "Select")
'  ï\é¶ ﬂ¿∞›æØƒ
'       Call .PatanSet( ﬂ¿∞›NO,ï\é¶ÉpÉ^Å[Éì)
        Call .PatanSet(0, "0,1,2,3,4,5")
        .InitDisp
    End With
End Sub
Private Sub Grid_RowChange1(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
    
'ÉeÉLÉXÉgÇ©ÇÁÉOÉäÉbÉhÇ…ëﬁî
    Call TextToVsGrid1(Val(iB1_MotoCD.Tag))
    gInt = vsGrid1.Row
    'Å™
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    'Å´
    If KeyCode = vbKeyDown Then
        If gInt + 1 > vsGrid1.Rows - 1 Then
            If Trim(vsGrid1.TextMatrix(vsGrid1.Rows - 1, 1)) = "" Then GoTo Grid_RowChange_End
            vsGrid1.Rows = vsGrid1.Rows + 1
        End If
        gInt = gInt + 1
    End If
    'PageUp
    If KeyCode = vbKeyPageUp Then
        If gInt - CisVsGrid1.DispMax < 1 Then
            gInt = 1
        Else
            gInt = gInt - CisVsGrid1.DispMax
        End If
    End If
    'PageDown
'    If KeyCode = vbKeyPageDown Then
'        If gInt + CisVsGrid.DispMax > Ttl_Kensu - 1 Then
'            gInt = Ttl_Kensu - 1
'        Else
'            gInt = gInt + CisVsGrid.DispMax
'        End If
'    End If
    vsGrid1.Row = gInt
'ÉOÉäÉbÉhÇ©ÇÁÉeÉLÉXÉgÇ…ï\é¶
    Call VsGridToText1(vsGrid1.Row)
    
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   ì¸óÕÉeÉLÉXÉgÇ©ÇÁÉOÉäÉbÉhÇ…à⁄Ç∑
'==========================================================='
Private Sub TextToVsGrid1(tRow As Integer)
    With vsGrid1
        .TextMatrix(tRow, 1) = iB1_MotoCD
        If Trim(iB1_MotoCD) = "" Then
            For gInt = 2 To .Cols - 1
                .TextMatrix(tRow, gInt) = ""
            Next gInt
            Exit Sub
        End If
        .TextMatrix(tRow, 0) = tRow
        .TextMatrix(tRow, 2) = iB1_MotoUke
        .TextMatrix(tRow, 3) = "Å@"
        .TextMatrix(tRow, 4) = iB1_ChgUke
        .TextMatrix(tRow, 5) = iB1_HenKeiyaku
    End With
End Sub
'==========================================================='
'   ÉOÉäÉbÉhÇ©ÇÁì¸óÕÉeÉLÉXÉgÇ…à⁄Ç∑ÅAì¸óÕÉeÉLÉXÉgÇÃà⁄ìÆ
'==========================================================='
Private Sub VsGridToText1(tRow As Integer)
    With vsGrid1
        'ïœä∑å≥CD
        iB1_MotoCD.Tag = tRow
        iB1_MotoCD = RTrim(.TextMatrix(tRow, 1))
        iB1_MotoCD.Move iB1_MotoCD.Left, vsGrid1.CellTop, iB1_MotoCD.Width, vsGrid1.CellHeight
        'éÛì¸
        iB1_MotoUke = RTrim(.TextMatrix(tRow, 2))
        iB1_MotoUke.Move iB1_MotoUke.Left, vsGrid1.CellTop, iB1_MotoUke.Width, vsGrid1.CellHeight
        'ïœä∑éÛì¸
        iB1_ChgUke = RTrim(.TextMatrix(tRow, 4))
        iB1_ChgUke.Move iB1_ChgUke.Left, vsGrid1.CellTop, iB1_ChgUke.Width, vsGrid.CellHeight
        'å_ñÒêÊ
        iB1_HenKeiyaku.Tag = tRow
        iB1_HenKeiyaku = RTrim(.TextMatrix(tRow, 5))
        iB1_HenKeiyaku.Move iB1_HenKeiyaku.Left, vsGrid1.CellTop, iB1_HenKeiyaku.Width, vsGrid1.CellHeight
            
        'ñÓàÛ
        B1lb_Migiya.Move B1lb_Migiya.Left, vsGrid1.CellTop, B1lb_Migiya.Width, vsGrid.CellHeight
    End With
End Sub
'******************************************'
'*****       ÉOÉäÉbÉhì‡óeÉZÉbÉg        *****
'******************************************'
Private Sub GridSet1()
    Dim ID          As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Long
' íäèoèàóù èâä˙æØƒ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' íäèoåèêî(çáåv)ëﬁî
' ∏ﬁÿØƒﬁçsçƒæØƒ & óÒà íuéÊìæ
    With vsGrid1
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid1.FixedGet(CisFun.Val2(ID))
        Next ID
        If Ttl_Kensu + 1 < CisVsGrid1.DispMax Then
            .Rows = CisVsGrid1.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not JTH_RDSTS
        With vsGrid1
            ID = ID + 1
            'áÇ
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            'ïœä∑å≥CD
            .TextMatrix(ID, GridCol(1)) = JTH.ïœä∑å≥CD
            'ïœä∑å≥éÛì¸
            .TextMatrix(ID, GridCol(2)) = JTH.ïœä∑å≥éÛì¸
            'ïœä∑êÊéÛì¸
            .TextMatrix(ID, GridCol(4)) = JTH.ïœä∑êÊéÛì¸
            'å_ñÒêÊ
            .TextMatrix(ID, GridCol(5)) = JTH.å_ñÒêÊ
        End With
        Call JTHReadNext(1)
    Loop
    
    With vsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub

Private Sub vsGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown Then Exit Sub
    Call TextToVsGrid1(Val(iB1_MotoCD.Tag))
    Call VsGridToText1(vsGrid1.Row)
End Sub
Private Sub vsGrid1_Click()
'    Call GridTextEnabled(True)
    Call TextToVsGrid1(Val(iB1_MotoCD.Tag))
    Call VsGridToText1(vsGrid1.Row)
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub

'===================================================================================== 2009/01/06 End

'===================================================================================== 2009/07/15 Start
'+==========================================================+
'+      çHíˆèÓïÒì¸óÕÉOÉäÉbÉhèàóù                      +
'+==========================================================+
'+----------------------------------------------------------+
'+   å©èoÇµæØƒ & èâä˙èÛë‘æØƒ
'+----------------------------------------------------------+
Private Sub GridInit2()
    With CisVsGrid2
'Å@∏ﬁÿØƒﬁèâä˙âªèàóù
        .Clear
'       Call .FixedSet(å©èoÇµà íuê›íË,óÒà íuê›íË,œ∞ºﬁèàóù,
'                      æŸåÖêî,óÒî‘çÜ,å©èoÇµì‡óe,ï™óﬁ¿≤Ãﬂ)
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 2, 0, "NO")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 4, 1, "ã@äB")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 10, 2, "ã@äBñº")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 6, 3, "ã@äBáÇ")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 10, 4, "ã@äBáÇñº")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 8, 5, "â“ì≠éûä‘")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 4, 6, "ë‰êî")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 6, 7, "ë„ã@äB")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 10, 8, "ë„ã@äBñº")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 8, 9, "ë„ã@äBáÇ")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 10, 10, "ë„ã@äBáÇñº")
'Å@∏ﬁÿØƒﬁëIëèàóùópÅ@ëﬁî¥ÿ±
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 0, 11, "Select")
'  ï\é¶ ﬂ¿∞›æØƒ
'       Call .PatanSet( ﬂ¿∞›NO,ï\é¶ÉpÉ^Å[Éì)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10")
        .InitDisp
    End With
End Sub
Private Sub Grid_RowChange2(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
    
'ÉeÉLÉXÉgÇ©ÇÁÉOÉäÉbÉhÇ…ëﬁî
    Call TextToVsGrid2(Val(iB1_Kikai.Tag))
    gInt = VsGrid2.Row
    'Å™
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid2_RowChange_End
        gInt = gInt - 1
    End If
    'Å´
    If KeyCode = vbKeyDown Then
        If gInt + 1 > VsGrid2.Rows - 1 Then
            If Trim(VsGrid2.TextMatrix(VsGrid2.Rows - 1, 1)) = "" And Trim(VsGrid2.TextMatrix(VsGrid2.Rows - 1, 3)) = "" Then GoTo Grid2_RowChange_End
            VsGrid2.Rows = VsGrid2.Rows + 1
        End If
        gInt = gInt + 1
    End If
    'PageUp
    If KeyCode = vbKeyPageUp Then
        If gInt - CisVsGrid2.DispMax < 1 Then
            gInt = 1
        Else
            gInt = gInt - CisVsGrid2.DispMax
        End If
    End If
    'PageDown
'    If KeyCode = vbKeyPageDown Then
'        If gInt + CisVsGrid.DispMax > Ttl_Kensu - 1 Then
'            gInt = Ttl_Kensu - 1
'        Else
'            gInt = gInt + CisVsGrid.DispMax
'        End If
'    End If
    VsGrid2.Row = gInt
'ÉOÉäÉbÉhÇ©ÇÁÉeÉLÉXÉgÇ…ï\é¶
    Call VsGridToText2(VsGrid2.Row)
    
Grid2_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   ì¸óÕÉeÉLÉXÉgÇ©ÇÁÉOÉäÉbÉhÇ…à⁄Ç∑
'==========================================================='
Private Sub TextToVsGrid2(tRow As Integer)
    With VsGrid2
        .TextMatrix(tRow, 1) = iB1_Kikai
        If Trim(iB1_Kikai) = "" And Trim(iB1_KikaiNo) = "" Then
            For gInt = 2 To .Cols - 1
                .TextMatrix(tRow, gInt) = ""
            Next gInt
            Exit Sub
        End If
        .TextMatrix(tRow, 0) = tRow
        .TextMatrix(tRow, 1) = iB1_Kikai
        .TextMatrix(tRow, 2) = RTrim(B1lb_KikaiNM)
        .TextMatrix(tRow, 3) = iB1_KikaiNo
        .TextMatrix(tRow, 4) = RTrim(B1lb_KikaiNoNM)
        If RTrim(iB1_KadoTime) <> "" Then
           .TextMatrix(tRow, 5) = Format(CisFun.Val2(iB1_KadoTime), "##0.00")
        Else
           .TextMatrix(tRow, 5) = ""
        End If
        .TextMatrix(tRow, 6) = Format(CisFun.Val2(iB1_Daisu), "###")
        .TextMatrix(tRow, 7) = iB1_DaiKikai
        .TextMatrix(tRow, 8) = RTrim(B1lb_DaiKikaiNM)
        .TextMatrix(tRow, 9) = iB1_DaiKikaiNo
        .TextMatrix(tRow, 10) = RTrim(B1lb_DaiKikaiNoNM)
    End With
End Sub
'==========================================================='
'   ÉOÉäÉbÉhÇ©ÇÁì¸óÕÉeÉLÉXÉgÇ…à⁄Ç∑ÅAì¸óÕÉeÉLÉXÉgÇÃà⁄ìÆ
'==========================================================='
Private Sub VsGridToText2(tRow As Integer)
    With VsGrid2
        'ã@äB
        iB1_Kikai.Tag = tRow
        iB1_Kikai = RTrim(.TextMatrix(tRow, 1))
        iB1_Kikai.Move iB1_Kikai.Left, VsGrid2.CellTop, iB1_Kikai.Width, VsGrid2.CellHeight
        'ã@äBñº
        B1lb_KikaiNM = RTrim(.TextMatrix(tRow, 2))
        B1lb_KikaiNM.Move B1lb_KikaiNM.Left, VsGrid2.CellTop, B1lb_KikaiNM.Width, VsGrid2.CellHeight
        'ã@äBáÇ
        iB1_KikaiNo = RTrim(.TextMatrix(tRow, 3))
        iB1_KikaiNo.Move iB1_KikaiNo.Left, VsGrid2.CellTop, iB1_KikaiNo.Width, VsGrid2.CellHeight
        'ã@äBáÇñº
        B1lb_KikaiNoNM = RTrim(.TextMatrix(tRow, 4))
        B1lb_KikaiNoNM.Move B1lb_KikaiNoNM.Left, VsGrid2.CellTop, B1lb_KikaiNoNM.Width, VsGrid2.CellHeight
        'â“ì≠éûä‘
        iB1_KadoTime = RTrim(.TextMatrix(tRow, 5))
        iB1_KadoTime.Move iB1_KadoTime.Left, VsGrid2.CellTop, iB1_KadoTime.Width, VsGrid2.CellHeight
        'ë‰êî
        iB1_Daisu = RTrim(.TextMatrix(tRow, 6))
        iB1_Daisu.Move iB1_Daisu.Left, VsGrid2.CellTop, iB1_Daisu.Width, VsGrid2.CellHeight
        'ë„ë÷ã@äB
        iB1_DaiKikai = RTrim(.TextMatrix(tRow, 7))
        iB1_DaiKikai.Move iB1_DaiKikai.Left, VsGrid2.CellTop, iB1_DaiKikai.Width, VsGrid2.CellHeight
        'ë„ë÷ã@äBñº
        B1lb_DaiKikaiNM = RTrim(.TextMatrix(tRow, 8))
        B1lb_DaiKikaiNM.Move B1lb_DaiKikaiNM.Left, VsGrid2.CellTop, B1lb_DaiKikaiNM.Width, VsGrid2.CellHeight
        'ë„ë÷ã@äBáÇ
        iB1_DaiKikaiNo = RTrim(.TextMatrix(tRow, 9))
        iB1_DaiKikaiNo.Move iB1_DaiKikaiNo.Left, VsGrid2.CellTop, iB1_DaiKikaiNo.Width, VsGrid2.CellHeight
        'ë„ë÷ã@äBáÇñº
        B1lb_DaiKikaiNoNM = RTrim(.TextMatrix(tRow, 10))
        B1lb_DaiKikaiNoNM.Move B1lb_DaiKikaiNoNM.Left, VsGrid2.CellTop, B1lb_DaiKikaiNoNM.Width, VsGrid2.CellHeight
            
    End With
End Sub
'******************************************'
'*****       ÉOÉäÉbÉhì‡óeÉZÉbÉg        *****
'******************************************'
Private Sub GridSet2()
    Dim ID          As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Long
' íäèoèàóù èâä˙æØƒ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' íäèoåèêî(çáåv)ëﬁî
' ∏ﬁÿØƒﬁçsçƒæØƒ & óÒà íuéÊìæ
    With VsGrid2
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid2.FixedGet(CisFun.Val2(ID))
        Next ID
        If Ttl_Kensu + 1 < CisVsGrid2.DispMax Then
            .Rows = CisVsGrid2.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not TKT_RDSTS
        With VsGrid2
            ID = ID + 1
            'áÇ
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            'ã@äB
            .TextMatrix(ID, GridCol(1)) = TKT.ã@äB
            'ã@äBñº
            .TextMatrix(ID, GridCol(2)) = TKT.ã@äBñº
            'ã@äBáÇ
            .TextMatrix(ID, GridCol(3)) = TKT.ã@äBNO
            'ã@äBáÇñº
            .TextMatrix(ID, GridCol(4)) = TKT.ã@äBNOñº
            'â“ì≠éûä‘
            .TextMatrix(ID, GridCol(5)) = Format(TKT.â“ì≠éûä‘, "##0.00")
            'ë‰êî
            .TextMatrix(ID, GridCol(6)) = Format(TKT.ë‰êî, "###")
            'ë„ë÷ã@äB
            .TextMatrix(ID, GridCol(7)) = TKT.ë„ë÷ã@äB
            'ë„ë÷ã@äBñº
            .TextMatrix(ID, GridCol(8)) = TKT.ë„ë÷ã@äBñº
            'ë„ë÷ã@äBáÇ
            .TextMatrix(ID, GridCol(9)) = TKT.ë„ë÷ã@äBNO
            'ë„ë÷ã@äBáÇñº
            .TextMatrix(ID, GridCol(10)) = TKT.ë„ë÷ã@äBNOñº
        End With
        Call TKTReadNext(1)
    Loop
    
    With VsGrid2
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    
End Sub

Private Sub vsGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown Then Exit Sub
    Call TextToVsGrid2(Val(iB1_Kikai.Tag))
    Call VsGridToText2(VsGrid2.Row)
End Sub
Private Sub vsGrid2_Click()
'    Call GridTextEnabled(True)
    Call TextToVsGrid2(Val(iB1_Kikai.Tag))
    Call VsGridToText2(VsGrid2.Row)
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub

'===================================================================================== 2009/07/15 End

