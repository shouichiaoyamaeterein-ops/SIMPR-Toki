VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{BA5D7819-2B67-11D4-806A-00000EA57F9E}#3.0#0"; "CISYMD~1.OCX"
Begin VB.Form CXU0520 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "îÑè„èÛãµä«óù"
   ClientHeight    =   11010
   ClientLeft      =   1140
   ClientTop       =   1140
   ClientWidth     =   15375
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z µ∞¿ﬁ∞
   ScaleHeight     =   11010
   ScaleWidth      =   15375
   WindowState     =   2  'ç≈ëÂâª
   Begin Cis3D_v60.CIS3D G2Color 
      Height          =   135
      Index           =   0
      Left            =   8760
      Top             =   420
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   238
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   8610
      Left            =   60
      ScaleHeight     =   8550
      ScaleWidth      =   15165
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2220
      Width           =   15225
      Begin vsFlexLib.vsFlexArray VSGrid1 
         Height          =   7005
         Left            =   30
         TabIndex        =   3
         Top             =   45
         Width           =   15090
         _Version        =   196608
         _ExtentX        =   26617
         _ExtentY        =   12356
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
         FormatString    =   $"CXU0520.frx":0000
         Rows            =   21
         Cols            =   13
         BackColor       =   12648447
         BackColorBkg    =   12632256
         HighLight       =   0
         RowHeightMin    =   330
         SelectionMode   =   1
         BackColorAlternate=   12648447
         Begin Cis3D_v60.CIS3D GTitle 
            Height          =   285
            Index           =   1
            Left            =   2385
            Top             =   30
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   503
            BackColor       =   -2147483633
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
            cBoderColor1    =   -2147483633
            cBoderColor2    =   -2147483633
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D GTitle 
            Height          =   285
            Index           =   2
            Left            =   3840
            Top             =   30
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   503
            BackColor       =   -2147483633
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
            cBoderColor1    =   -2147483633
            cBoderColor2    =   -2147483633
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D GTitle 
            Height          =   285
            Index           =   3
            Left            =   5295
            Top             =   30
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
            BackColor       =   -2147483633
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
            cBoderColor1    =   -2147483633
            cBoderColor2    =   -2147483633
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D GTitleDay 
            Height          =   285
            Index           =   0
            Left            =   6015
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   503
            BackColor       =   -2147483633
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
            cBoderColor1    =   -2147483633
            cBoderColor2    =   -2147483633
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D GTitleDay 
            Height          =   285
            Index           =   1
            Left            =   7260
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   503
            BackColor       =   -2147483633
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
            cBoderColor1    =   -2147483633
            cBoderColor2    =   -2147483633
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D GTitleDay 
            Height          =   285
            Index           =   2
            Left            =   8505
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   503
            BackColor       =   -2147483633
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
            cBoderColor1    =   -2147483633
            cBoderColor2    =   -2147483633
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D GTitleDay 
            Height          =   285
            Index           =   3
            Left            =   9750
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   503
            BackColor       =   -2147483633
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
            cBoderColor1    =   -2147483633
            cBoderColor2    =   -2147483633
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D GTitleDay 
            Height          =   285
            Index           =   5
            Left            =   12240
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   503
            BackColor       =   -2147483633
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
            cBoderColor1    =   -2147483633
            cBoderColor2    =   -2147483633
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D GTitleDay 
            Height          =   285
            Index           =   4
            Left            =   10995
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   503
            BackColor       =   -2147483633
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
            cBoderColor1    =   -2147483633
            cBoderColor2    =   -2147483633
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   2
         End
      End
      Begin vsFlexLib.vsFlexArray VsGrid2 
         Height          =   1395
         Left            =   15
         TabIndex        =   8
         Top             =   7110
         Width           =   15090
         _Version        =   196608
         _ExtentX        =   26617
         _ExtentY        =   2461
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
         FormatString    =   $"CXU0520.frx":0091
         Rows            =   4
         Cols            =   11
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   12648447
         BackColorBkg    =   12632256
         HighLight       =   0
         RowHeightMin    =   330
         SelectionMode   =   1
         BackColorAlternate=   12648447
      End
   End
   Begin VB.PictureBox H1_Area1 
      Height          =   1290
      Left            =   60
      ScaleHeight     =   1230
      ScaleWidth      =   8310
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   825
      Width           =   8370
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1140
         Left            =   6285
         Top             =   45
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   2011
         ForeColor       =   16711680
         Caption         =   "é©é–çHãÊ"
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
         cPositionY      =   60
         Begin VB.ComboBox H1Comb_Jisya 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   270
            Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
            TabIndex        =   2
            Top             =   510
            Width           =   1395
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   1140
         Left            =   2520
         Top             =   45
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   2011
         ForeColor       =   16711680
         Caption         =   "å_ñÒêÊ"
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
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   150
            TabIndex        =   1
            Top             =   510
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
            Left            =   1140
            Top             =   510
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
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   1140
         Index           =   0
         Left            =   45
         Top             =   45
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   2011
         ForeColor       =   16711680
         Caption         =   "éwíËì˙"
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
         cPositionY      =   60
         Begin CisYMD_V60.CisYMD iH1_YMD 
            Height          =   360
            Left            =   330
            TabIndex        =   0
            Top             =   510
            Width           =   1770
            _ExtentX        =   3069
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
            Object.Height          =   360
            Object.Width           =   1770
            cSize           =   -1  'True
            cChkResult      =   0   'False
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
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "Å@Å@Å@Åy îÑè„èÛãµä«óù Åz"
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
      Begin Cis3D_v60.CIS3D UC_3D32 
         Height          =   255
         Left            =   8880
         Top             =   45
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   450
         BackColor       =   14737632
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
         cBoderColor1    =   14737632
         cBoderColor2    =   14737632
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
            BackColor       =   14737632
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
            cBoderColor1    =   14737632
            cBoderColor2    =   14737632
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
         Left            =   11925
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
   Begin VB.TextBox Dummy 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   345
      IMEMode         =   3  'µÃå≈íË
      Left            =   510
      TabIndex        =   4
      Text            =   "D"
      Top             =   900
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Left            =   780
      Top             =   900
   End
   Begin VB.Timer Timer2 
      Left            =   1500
      Top             =   900
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   705
      Left            =   10875
      ScaleHeight     =   645
      ScaleWidth      =   4365
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   810
      Width           =   4425
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   585
         Left            =   3450
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
         Left            =   2610
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
         Left            =   1770
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   60
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
      Begin CisBtn_60.CisBtn PB_PRT 
         Height          =   585
         Left            =   915
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
         BtnType         =   7
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
   End
   Begin Cis3D_v60.CIS3D G2Color 
      Height          =   135
      Index           =   1
      Left            =   8760
      Top             =   570
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   238
      BackColor       =   16777152
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
   End
   Begin Cis3D_v60.CIS3D G2Color 
      Height          =   135
      Index           =   2
      Left            =   8760
      Top             =   720
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   238
      BackColor       =   16761024
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
   End
   Begin Cis3D_v60.CIS3D G2Color 
      Height          =   135
      Index           =   3
      Left            =   8760
      Top             =   870
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   238
      BackColor       =   12640511
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
   End
   Begin Cis3D_v60.CIS3D G2Color 
      Height          =   135
      Index           =   4
      Left            =   8760
      Top             =   1020
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   238
      BackColor       =   12648447
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
   End
   Begin Cis3D_v60.CIS3D G2Color 
      Height          =   135
      Index           =   5
      Left            =   8760
      Top             =   1170
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   238
      BackColor       =   12648384
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
   End
   Begin Cis3D_v60.CIS3D G2Color 
      Height          =   135
      Index           =   6
      Left            =   8760
      Top             =   1320
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   238
      BackColor       =   8438015
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
   End
   Begin Cis3D_v60.CIS3D G2Color 
      Height          =   135
      Index           =   7
      Left            =   8760
      Top             =   1470
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   238
      BackColor       =   8454143
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
   End
   Begin Cis3D_v60.CIS3D G2Color 
      Height          =   135
      Index           =   8
      Left            =   8760
      Top             =   1620
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   238
      BackColor       =   8454016
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
   End
   Begin Cis3D_v60.CIS3D G2Color 
      Height          =   135
      Index           =   9
      Left            =   8760
      Top             =   1770
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   238
      BackColor       =   16777088
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
   End
End
Attribute VB_Name = "CXU0520"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  îÑè„é¿ê—ñ‚çáÇπ
'**       ÉtÉHÅ[ÉÄID    :  CXU0410
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2008/05/19  By CIS
'**       ïœ  çX  ì˙    :  2008/06/12 By CIS é¿ê—îNåéÇåüé˚ì˙/èoâ◊ì˙Ç…èCê≥
'**                                          éÊà¯êÊÇÃåvéZãÊï™ëŒâû
'**
'**       ïœ  çX  ì˙    :  2009/02/10 By CIS
'**       ïœ  çX  ì˙    :  2009/02/12 By CIS
'**       ïœ  çX  ì˙    :  2009/04/01 By CIS ÅyèCê≥Åzåéå◊Ç¨ÉfÅ[É^ÇÃèÍçáéwíËåéï™ÇµÇ©èoóÕÇ≥ÇÍÇ»Ç¢
'**                                                  éwíË)2009/03 ÅÀ 04/01 ÇÃÉfÅ[É^Ç™ 03/01 ÇÃÉfÅ[É^Ç…
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim ProcHB          As String * 2   ' ì¸óÕàÊêÿë÷
    Dim FormAct         As Boolean      ' Ã´∞—±∏√®Ãﬁ Ã◊∏ﬁ
    Dim wObj            As Object       ' µÃﬁºﬁ™∏ƒ‹∞∏¥ÿ±
    Dim Sl_Where        As String       ' SQLèåè
    Dim Ttl_Kensu       As Integer      ' ï\é¶åèêî
    Dim mDispMax        As Integer      ' ∏ﬁÿØƒﬁëSçsêîëﬁî]
    
    
    Dim CisComb1        As Cis_Comb.CisCombo
    
    Dim CisVsGrid       As New CisVsGrid3
    Dim CisVsGrid2      As New CisVsGrid3
    
    Dim UKB_IRN         As Integer
    
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås    ÇkÇèÇÅÇÑÉCÉxÉìÉg    Åt              +
'+----------------------------------------------------------+
Private Sub Form_Load()

    CisDB.TimeOut = 3000
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
'   *-------------------*
'   *   ∏ﬁÿØƒﬁèâä˙èàóù
'   *-------------------*
    With CisVsGrid
        Set .GridObj = VSGrid1
        .NoSpaceCol = 1
        .NumberSet = No
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 12
        .InitGet = False
        .Init
    End With
    
    Call GridInit
'   *-------------------*
'   *   ∏ﬁÿØƒﬁèâä˙èàóù2
'   *-------------------*
    With CisVsGrid2
        Set .GridObj = VsGrid2
        .NoSpaceCol = 1
        .NumberSet = No
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 11
        .InitGet = False
        .Init
    End With
            
'   +------------------+
'   +   ∫›ŒﬁŒﬁØ∏ΩæØƒ   +
'   +------------------+
    Set CisComb1 = New CisCombo
    With CisComb1
        Set .Connect = CisDB
        Set .CombObj = H1Comb_Jisya
        .Control = "ñºèÃÉ}ÉXÉ^"
        .ComboName = "é©é–çHãÊ"
        .Ess = False
        .IniPath = gIniExe
        .SetSpace = 50
        .ComboInit
    End With
    
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
    
    mDispMax = CisVsGrid.DispMax
    lb_Kensu = ""

    Dummy.Left = -1000
    Dummy.Enabled = False
    Timer1.Interval = 1000   ' éûçèï\é¶ä‘äu
    Timer1.Enabled = True    ' éûçèï\é¶äJén
'
'   #------------------#
'   #    èâä˙ì¸óÕê›íË   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   # âÊñ íÜâõà íuï\é¶   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    If Rv_Call3 = "" Then
       Me.Show
    End If
End Sub
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    ÅsÇ`ÇÉÇîÇâÇñÇÅÇîÇÖ ÉCÉxÉìÉgÅt              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    gSL_Select = ""
    gSL_Select = gSL_Select & " SELECT * FROM IRNAKSIC"
    gSL_Select = gSL_Select & "  WHERE RNO = 50"
    gSL_Select = gSL_Select & "    AND BK = 'Option'"
    If IRNRead(gSL_Select, 1) Then
       UKB_IRN = CisFun.Val2(IRN.oyian)
    End If
    If UKB_IRN = 0 Then
        gSL_Select = "SELECT MAX(é¿ê—îNåé) é¿ê—îNåé FROM åüé˚ÉeÅ[ÉuÉã"
    Else
        gSL_Select = "SELECT MAX(é¿ê—îNåé) é¿ê—îNåé FROM èoâ◊ÉeÅ[ÉuÉã"
    End If
    
    KojoDsp
    
    Call CisComb1.ComboSet
    
    If Rv_Call3 <> "" Then
       If Format(Date, "MM") = Mid(RV_YM, 5, 2) Then
          iH1_YMD = Format(Date, "yyyyMMdd")
       Else
          iH1_YMD = RV_YM & "01"
       End If
       iH1_Torcd = RV_KTorcd
       Call iH1_Torcd_LostFocus
       If RV_Code <> "" Then
          CisComb1.Code = RV_Code
          Call CisComb1.TextGet
       End If
       Call ReturnPress
    Else
       iH1_YMD = Format(Date, "yyyyMMdd")
       TitleDsp
    End If
    
    
        
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås QueryUnloadÉCÉxÉìÉg Åt                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås ÇjÇÖÇôÇcÇèÇóÇéÉCÉxÉìÉg Åt               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄÅs ÉtÉ@ÉìÉNÉVÉáÉìÉLÅ[ópÉAÉNÉZÉâÉåÅ[É^ Åt       +
'+----------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Key_Acc = False
    If Not FormAct Then Exit Function
    If UCase(IVENT) <> "DOWN" Then Exit Function
    
    Select Case Key_Code
        Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  'ÅyåüçıÅz
        Case vbKeyF8:       If PB_PRT.Visible Then Call PB_PRT_Click    'ÅyàÛç¸Åz
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    'ÅyèIóπÅz
        Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    'ÅyéÊè¡Åz
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
        Me.Hide
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
        Call HeadClear
        iH1_YMD.SetFocus
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B1" Then
        lb_Kensu = ""
        CisVsGrid.DispMax = mDispMax
        VSGrid1.Rows = mDispMax
        CisVsGrid.Clear
        KojoDsp
        Call GridInit
        ProcHB = "H1"
        Call DispChange(ProcHB)
        iH1_YMD.SetFocus
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub

'+----------------------------+
'+    àÛç¸   (F8)             +
'+----------------------------+
Private Sub PB_PRT_Click()
    CXU0520_YM = CisFun.Mid2(iH1_YMD, 1, 6)
    CXU0520_MD1 = GTitleDay(0)
    CXU0520_MD2 = GTitleDay(1)
    CXU0520_MD3 = GTitleDay(2)
    CXU0520_MD4 = GTitleDay(3)
    CXU0520_MD5 = GTitleDay(4)
    CXU0520_MD6 = GTitleDay(5)
    
    CXU0525.Show vbModal
    Unload CXU0525
    Set CXU0525 = Nothing
    
End Sub

'+----------------------------+
'+    åüçı   (F7)             +
'+----------------------------+
Private Sub PB_Look_Click()
    If ProcHB = "H1" Then
       If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    End If
End Sub
'+----------------------------+
'+      åü         çı         +
'+----------------------------+
Private Sub Look_Torcd()

    RV_TorKb = 2    ' ìæà”êÊ
    RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    iH1_Torcd.SetFocus

    If RV_Rtn Then
       iH1_Torcd = RV_TorcdK
       Call iH1_Torcd_LostFocus
    End If

End Sub
'+----------------------------------------+
'+  iH1_Torcd ( GotFocus )   å_ñÒêÊ
'+----------------------------------------+
Private Sub iH1_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
'+----------------------------------------+
'+  iH1_Torcd ( LostFocus )  å_ñÒêÊ
'+----------------------------------------+
Private Sub iH1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
       PB_Look.Visible = False
       PB_Look.Tag = ""
    End If
    If TorNmGet(iH1_Torcd, 2) Then
        H1lb_Torcd = TRM_RName
        If Trim(H1lb_Torcd) = "" Then
           H1lb_Torcd = TRM_Name
        End If
    Else
        H1lb_Torcd = ""
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

'    PB_CAN.Visible = Not H1Mode
    PB_Ent.Visible = H1Mode
    PB_PRT.Visible = B1Mode
' ¿ﬁ–∞çÄñ⁄(ì¸óÕã÷é~)
    Dummy.Enabled = False
    
    If Rv_Call3 <> "" Then
        PB_END.BtnCaption = "ñﬂÇÈ"
    End If
'
End Function
'****************************
'*      ì¸óÕèàóù(Enter)      *
'****************************
Private Sub ReturnPress()
   Call CisFrm.MousePT(11)
   If ProcHB = "H1" Then
      ProcHB = "B1"
      If Not Head1Chk Then
          ProcHB = "H1"
          GoTo ReturnPress_Ed
      End If
      Call DispChange(ProcHB)
      VSGrid1.SetFocus
      GoTo ReturnPress_Ed
   End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1èâä˙âª
'****************************
Private Sub HeadClear()
    For Each wObj In Me.Controls
        If wObj.Name Like "iH*" Or _
           wObj.Name Like "H1lb_*" Then wObj = ""
        If wObj.Name Like "Ch_*" Then wObj.Value = 0
    Next wObj
End Sub

'****************************
'*      HEAD1¡™Ø∏èàóù
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
 
' éwíËì˙
    If Trim(iH1_YMD) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "     éwíËì˙Çì¸óÕÇµÇƒâ∫Ç≥Ç¢      "
            .MB_MSG(3) = "   "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YMD.SetFocus
        Exit Function
    End If
    If Not IsDate(iH1_YMD.cYear & "/" & iH1_YMD.cMonth & "/01") Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "     éwíËì˙Çê≥ÇµÇ≠ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢       "
            .MB_MSG(3) = "   "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YMD.SetFocus
        Exit Function
    End If

    If Not DBInput Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    äYìñÉfÅ[É^Ç™ë∂ç›ÇµÇ‹ÇπÇÒ°            "
            .MB_MSG(3) = "    çƒìxíäèoèåèÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢        "
            .MB_Button = Error
            .MB_Title = "íäèoèàóù"
            Call .MBOX
       End With
       Exit Function
    End If
    
    Call DBInput2
    
    Head1Chk = True
End Function
'*******************************
'*      É^ÉCÉgÉãÉoÅ[ï“èW        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'+==========================================================+
'+      ÉOÉäÉbÉhèàóù                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   å©èoÇµæØƒ & èâä˙èÛë‘æØƒ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'Å@∏ﬁÿØƒﬁèâä˙âªèàóù
'        .Clear ("NO")
'       Call .FixedSet(å©èoÇµà íuê›íË,óÒà íuê›íË,œ∞ºﬁèàóù,
'                      æŸåÖêî,óÒî‘çÜ,å©èoÇµì‡óe)
        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 1, 0, "")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 20, 1, "å_ñÒêÊ")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 13, 2, "åéìxó\ë™")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 13, 3, "åéìxé¿ê—")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 6, 4, "êiíª")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 11, 5, "05/01")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 11, 6, "05/02")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 11, 7, "05/03")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 11, 8, "05/04")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 11, 9, "05/05")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 11, 10, "05/06")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç»Çµ, 11, 11, "ì˙ìñÇË")
        
'Å@∏ﬁÿØƒﬁëIëèàóùópÅ@ëﬁî¥ÿ±
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 0, 12, "Select")
'  ï\é¶ ﬂ¿∞›æØƒ
'       Call .PatanSet( ﬂ¿∞›NO,ï\é¶ÉpÉ^Å[Éì)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11")
        .InitDisp
    End With
End Sub
'****************************
'*      É{ÉfÉBì‡óeÉZÉbÉg
'****************************
Private Function BodySet() As Boolean
    Dim ID          As Long
    Dim ID2         As Long
    Dim wChk        As Boolean
    Dim GridCol()   As Long

    BodySet = False

    On Error GoTo BodySet_Err:
        
    Ttl_Kensu = CisDB.RecordCount
    lb_Kensu.Caption = Format(Ttl_Kensu, "#,###,##0")

    DoEvents

' ∏ﬁÿØƒﬁçsçƒæØƒ & óÒà íuéÊìæ
    With VSGrid1
        .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(ID)
        Next ID
    End With
    ID = 1
    If CisDB.RecordCount + 1 > VSGrid1.Rows Then
       VSGrid1.Rows = CisDB.RecordCount + 1
    End If

    Do Until Not USV_RDSTS
        With VSGrid1
            ' å_ñÒêÊ
            .TextMatrix(ID, GridCol(1)) = RTrim(USV.ï“èWå_ñÒêÊ)
            ' ó\ë™ã‡äz
            .TextMatrix(ID, GridCol(2)) = Format(USV.ó\ë™ã‡äz, "#,###")
            ' é¿ê—ã‡äz
            .TextMatrix(ID, GridCol(3)) = Format(USV.é¿ê—ã‡äz, "#,###")
            If CisFun.Val2(USV.êiíª) <> 0 Then
               .TextMatrix(ID, GridCol(4)) = Format(USV.êiíª, "#,##0.0;;#") & "%"
            End If
            .TextMatrix(ID, GridCol(5)) = Format(USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2))), "#,###")
            .TextMatrix(ID, GridCol(6)) = Format(USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2))), "#,###")
            .TextMatrix(ID, GridCol(7)) = Format(USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2))), "#,###")
            .TextMatrix(ID, GridCol(8)) = Format(USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2))), "#,###")
            .TextMatrix(ID, GridCol(9)) = Format(USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2))), "#,###")
            .TextMatrix(ID, GridCol(10)) = Format(USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2))), "#,###")
            .TextMatrix(ID, GridCol(11)) = Format(USV.ì˙ìñÇËêî, "#,###")
            ID = ID + 1
        End With
        Call USVReadNext
    Loop
    Call USVClose

    CisVsGrid.DispMax = VSGrid1.Rows

BodySet_ED:

    Call GridInit
    With VSGrid1
         .Col = 1: .ColSel = .Cols - 1
         .Row = 1: .Redraw = True
    End With
    BodySet = True

BodySet_ED1:
    Exit Function
BodySet_Err:
     CisFun.ErrorBox ("BodySet")
End Function

'****************************
'*      É{ÉfÉBì‡óeÉZÉbÉg
'****************************
Private Function BodySet2() As Boolean
    Dim ID          As Long
    Dim ID2         As Long
    Dim wChk        As Boolean
    Dim GridCol()   As Long
    
    Dim GKeiYs      As Currency     ' ó\ë™ã‡äz
    Dim GKeiJs      As Currency     ' é¿ê—ã‡äz
    Dim GKeiMD1     As Currency     ' XX/XX1
    Dim GKeiMD2     As Currency     ' XX/XX2
    Dim GKeiMD3     As Currency     ' XX/XX3
    Dim GKeiMD4     As Currency     ' XX/XX4
    Dim GKeiMD5     As Currency     ' XX/XX5
    Dim GKeiMD6     As Currency     ' XX/XX6
    Dim GKeiDA      As Currency     ' ì˙ìñÇË
    
    GKeiYs = 0: GKeiJs = 0
    GKeiMD1 = 0: GKeiMD2 = 0: GKeiMD3 = 0
    GKeiMD4 = 0: GKeiMD5 = 0: GKeiMD6 = 0
    GKeiDA = 0

    BodySet2 = False

    On Error GoTo BodySet2_Err:
        
    Ttl_Kensu = CisDB.RecordCount
    lb_Kensu.Caption = Format(Ttl_Kensu, "#,###,##0")

    DoEvents

' ∏ﬁÿØƒﬁçsçƒæØƒ & óÒà íuéÊìæ
    With VsGrid2
        .Redraw = False
        ReDim GridCol(.Cols)
    End With
    ID = 1
    If CisDB.RecordCount + 1 > VsGrid2.Rows Then
       VsGrid2.Rows = CisDB.RecordCount + 1
    End If

    Do Until Not USV_RDSTS
        With VsGrid2
            ' çsê›íË
            ID = CisFun.Val2(USV.é©é–çHãÊ) + 1
            ' ó\ë™ã‡äz
            .TextMatrix(ID, 1) = CisFun.RSetFld(USV.ó\ë™ã‡äz, 13, "#,###")
            GKeiYs = GKeiYs + USV.ó\ë™ã‡äz
            ' é¿ê—ã‡äz
            .TextMatrix(ID, 2) = CisFun.RSetFld(USV.é¿ê—ã‡äz, 13, "#,###")
            GKeiJs = GKeiJs + USV.é¿ê—ã‡äz
            If CisFun.Val2(USV.êiíª) <> 0 Then
               .TextMatrix(ID, 3) = CisFun.RSetFld(USV.êiíª, 5, "#,##0.0;;#") & "%"
            End If
            .TextMatrix(ID, 4) = CisFun.RSetFld(USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2))), 11, "#,###")
            GKeiMD1 = GKeiMD1 + USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)))
            .TextMatrix(ID, 5) = CisFun.RSetFld(USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2))), 11, "#,###")
            'GKeiMD2 = GKeiMD2 + USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)))    2009/02/10
            GKeiMD2 = GKeiMD2 + USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)))
            .TextMatrix(ID, 6) = CisFun.RSetFld(USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2))), 11, "#,###")
            'GKeiMD3 = GKeiMD3 + USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)))    2009/02/10
            GKeiMD3 = GKeiMD3 + USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)))
            .TextMatrix(ID, 7) = CisFun.RSetFld(USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2))), 11, "#,###")
            'GKeiMD4 = GKeiMD4 + USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)))    2009/02/10
            GKeiMD4 = GKeiMD4 + USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)))
            .TextMatrix(ID, 8) = CisFun.RSetFld(USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2))), 11, "#,###")
            'GKeiMD5 = GKeiMD5 + USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)))    2009/02/10
            GKeiMD5 = GKeiMD5 + USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)))
            .TextMatrix(ID, 9) = CisFun.RSetFld(USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2))), 11, "#,###")
            'GKeiMD6 = GKeiMD6 + USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)))    2009/02/10
            GKeiMD6 = GKeiMD6 + USV.é¿ê—ã‡äzT(CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)))
            .TextMatrix(ID, 10) = CisFun.RSetFld(USV.ì˙ìñÇËêî, 11, "#,###")
            GKeiDA = GKeiDA + USV.ì˙ìñÇËêî
        End With
        Call USVReadNext
    Loop
    
    With VsGrid2
        .TextMatrix(0, 1) = CisFun.RSetFld(GKeiYs, 13, "#,###")
        .TextMatrix(0, 2) = CisFun.RSetFld(GKeiJs, 13, "#,###")
        If GKeiJs <> 0 And GKeiYs <> 0 Then
           .TextMatrix(0, 3) = CisFun.RSetFld(CisFun.HasuProc(GKeiJs / GKeiYs * 100, éléÃå‹ì¸, 1), 5, "#,##0.0;;#") & "%"
        End If
        .TextMatrix(0, 4) = CisFun.RSetFld(GKeiMD1, 11, "#,###")
        .TextMatrix(0, 5) = CisFun.RSetFld(GKeiMD2, 11, "#,###")
        .TextMatrix(0, 6) = CisFun.RSetFld(GKeiMD3, 11, "#,###")
        .TextMatrix(0, 7) = CisFun.RSetFld(GKeiMD4, 11, "#,###")
        .TextMatrix(0, 8) = CisFun.RSetFld(GKeiMD5, 11, "#,###")
        .TextMatrix(0, 9) = CisFun.RSetFld(GKeiMD6, 11, "#,###")
        .TextMatrix(0, 10) = CisFun.RSetFld(GKeiDA, 11, "#,###")
    End With
    
    Call USVClose

BodySet2_ED:

    With VsGrid2
         .Col = 1: .ColSel = .Cols - 1
         .Row = 1: .Redraw = True
    End With
    BodySet2 = True

BodySet2_ED1:
    Exit Function
BodySet2_Err:
     CisFun.ErrorBox ("BodySet2")
End Function

'+--------------------------+
'+     ÉfÅ[É^ÉxÅ[ÉXï\é¶     +
'+--------------------------+
Private Function DBInput() As Boolean
    Dim wCnt As Integer
    Dim wStr        As String
    Dim XStr        As String
    Dim wWhereS     As String   ' èoâ◊/åüé˚ÉeÅ[ÉuÉãóp WhereãÂ
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/04/01 INSERT START
    Dim wWhereSS    As String
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/04/01 INSERT END
    Dim wWhereN     As String   ' ì‡é¶ÉeÅ[ÉuÉãóp WhereãÂ
    Dim wKadou      As Integer
    
    DBInput = False
    
    Call TitleDsp
    
    Sl_Where = "": gSL_Select = ""
    
    With CisDB
        '+---------- ÉèÅ[ÉNÇÃçÌèú -----------+
        .SQL = ""
        .SQL = .SQL & " DELETE FROM ì‡é¶ì˙ïtï ÉèÅ[ÉN"
        .SQL = .SQL & "  WHERE èàóùí[ññ = SUBSTRING(HOST_NAME(), 1, 20)"
        Call CisDB.DBExec
    
        '+----- ÉèÅ[ÉNÇ…ÉfÅ[É^çÏê¨ -------
        For wCnt = 1 To 31
            wStr = Format(wCnt, "00")
            .SQL = ""
            .SQL = .SQL & " INSERT INTO ì‡é¶ì˙ïtï ÉèÅ[ÉN"
            .SQL = .SQL & "             (èàóùí[ññ, ì‡é¶îNåé, ïiî‘, î[ì¸êÊ, éÛì¸, ì‡é¶ãÊï™, î[ä˙, êîó )"
            .SQL = .SQL & " SELECT SUBSTRING(HOST_NAME(), 1, 20), ì‡é¶îNåé, ïiî‘, î[ì¸êÊ, éÛì¸,"
            .SQL = .SQL & "        ì‡é¶ãÊï™, ì‡é¶îNåé + '" & wStr & "', SUM(ì‡é¶êîó " & wStr & ")"
            .SQL = .SQL & "   FROM ì‡é¶ÉeÅ[ÉuÉã"
            .SQL = .SQL & "  WHERE ì‡é¶îNåé = '" & CisFun.Mid2(iH1_YMD, 1, 6) & "'"
            .SQL = .SQL & "  GROUP BY ì‡é¶îNåé, ïiî‘, î[ì¸êÊ, éÛì¸, ì‡é¶ãÊï™"
            Call CisDB.DBExec
        Next wCnt
    End With
    
' ï\é¶èåè ÉZÉbÉg
    With CisFun
        '+-------- èoâ◊ & ì‡é¶ÉfÅ[É^ --------+
        .Where_CreateKB = NewDocument
        ' éwíËì˙
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/12 UPDATE START
        '.Where_Create AlphaMe, "KS.é¿ê—îNåé", Mid(iH1_YMD, 1, 6)
        .Where_Create AlphaMe, "îNåé", Mid(iH1_YMD, 1, 6)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/12 UPDATE END
        ' å_ñÒêÊ
        .Where_Create AlphaMe, "å_ñÒêÊ", iH1_Torcd
         ' é©é–çHãÊ
         Call CisComb1.CodeGet
         .Where_Create AlphaMe, "é©é–çHãÊ", CisComb1.Code
         
         wWhereS = CisFun.Where_Phrase
         
    End With

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/12 INSERT START
    '+----------- ì‡é¶ÉfÅ[É^ÇÃÇ› -------------+
    With CisFun
        .Where_CreateKB = NewDocument
         ' îNåé
        .Where_Create AlphaMe, "îNåé", CisFun.Mid2(iH1_YMD, 1, 6)
         ' å_ñÒêÊ
         .Where_Create AlphaMe, "å_ñÒêÊ", iH1_Torcd
         ' é©é–çHãÊ
         Call CisComb1.CodeGet
         .Where_Create AlphaMe, "é©é–çHãÊ", CisComb1.Code
'         ' èoâ◊Ç»Çµ
'         .Where_Create Numeric, "èoâ◊éwé¶èëNO", 0, , Create
         
         wWhereN = CisFun.Where_Phrase
    End With
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/04/01 INSERT START
' ï\é¶èåè ÉZÉbÉg
    If CisFun.Val2(CisFun.Mid2(iH1_YMD, 7, 2)) < 7 Or CisFun.Val2(CisFun.Mid2(iH1_YMD, 7, 2)) > 25 Then
        
        CisFun.DateE_BefVal = iH1_YMD
        CisFun.DateEdit (Slash)
        If CisFun.Val2(CisFun.Mid2(iH1_YMD, 7, 2)) < 7 Then
           Call CisFun.Date_Add(iH1_YMD, M, -1)
        Else
           Call CisFun.Date_Add(iH1_YMD, D, 1)
        End If
        
        With CisFun
            '+-------- èoâ◊ & ì‡é¶ÉfÅ[É^ --------+
            .Where_CreateKB = NewDocument
    
            ' éwíËì˙
            .Where_Create AlphaMe, "îNåé", Mid(Replace(.Date_Result, "/", ""), 1, 6)
            ' å_ñÒêÊ
            .Where_Create AlphaMe, "å_ñÒêÊ", iH1_Torcd
             ' é©é–çHãÊ
             Call CisComb1.CodeGet
             .Where_Create AlphaMe, "é©é–çHãÊ", CisComb1.Code
             
             wWhereSS = CisFun.Where_Phrase
             
        End With
    End If
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/04/01 INSERT END

    
    '+------- â“ì≠ì˙êîéZèo ------+
    gSL_Select = ""
    gSL_Select = gSL_Select & " SELECT â“ìÆì˙êî FROM ÉJÉåÉìÉ_É}ÉXÉ^"
    gSL_Select = gSL_Select & "  WHERE éÊà¯êÊ = ''"
    gSL_Select = gSL_Select & "    AND éÊà¯êÊãÊï™ = 0"
    gSL_Select = gSL_Select & "    AND îN = '" & CisFun.Mid2(iH1_YMD, 1, 4) & "'"
    gSL_Select = gSL_Select & "    AND åé = '" & CisFun.Mid2(iH1_YMD, 5, 2) & "'"
    If CLMRead(gSL_Select) Then
       wKadou = CLM.â“ìÆì˙êî
    Else
       wKadou = 0
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/12 INSERT END
    gSL_Select = " SELECT å_ñÒêÊ, å_ñÒêÊñº, ï“èWå_ñÒêÊ, SUM(ó\ë™ã‡äz) ó\ë™ã‡äz, SUM(é¿ê—ã‡äz) é¿ê—ã‡äz," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "," & vbCrLf
    gSL_Select = gSL_Select & "        CASE WHEN SUM(ó\ë™ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
    gSL_Select = gSL_Select & "             CASE WHEN SUM(é¿ê—ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
    gSL_Select = gSL_Select & "                  ROUND(SUM(é¿ê—ã‡äz) / SUM(ó\ë™ã‡äz) * 100, 1)" & vbCrLf
    gSL_Select = gSL_Select & "        END  END êiíª," & vbCrLf
    If wKadou = 0 Then
       gSL_Select = gSL_Select & "        0 ì˙ìñÇËêî"
    Else
       gSL_Select = gSL_Select & "        ROUND(SUM(ó\ë™ã‡äz) / " & wKadou & ",1) ì˙ìñÇËêî"
    End If
    gSL_Select = gSL_Select & "   FROM ("
        
    gSL_Select = gSL_Select & " SELECT å_ñÒêÊ, å_ñÒêÊñº, ï“èWå_ñÒêÊ, 0 ó\ë™ã‡äz, SUM(é¿ê—ã‡äz) é¿ê—ã‡äz," & vbCrLf
    If CisFun.Mid2(GTitleDay(0), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
    Else
       gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(1), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
    Else
       gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(2), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
    Else
       gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(3), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
    Else
       gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(4), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
    Else
       gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(5), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "" & vbCrLf
    Else
       gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "" & vbCrLf
    End If
    If UKB_IRN = 0 Then
       gSL_Select = gSL_Select & "   FROM îÑè„åüé˚èÛãµä«óùÉrÉÖÅ[" & vbCrLf
    Else
       gSL_Select = gSL_Select & "   FROM îÑè„èoâ◊èÛãµä«óùÉrÉÖÅ[" & vbCrLf
    End If
    gSL_Select = gSL_Select & wWhereS & vbCrLf
    gSL_Select = gSL_Select & "  GROUP BY å_ñÒêÊ, å_ñÒêÊñº, ï“èWå_ñÒêÊ" & vbCrLf
    gSL_Select = gSL_Select & " UNION" & vbCrLf
    
    If CisFun.Val2(CisFun.Mid2(iH1_YMD, 7, 2)) < 7 Or CisFun.Val2(CisFun.Mid2(iH1_YMD, 7, 2)) > 25 Then
        gSL_Select = gSL_Select & " SELECT å_ñÒêÊ, å_ñÒêÊñº, ï“èWå_ñÒêÊ, 0 ó\ë™ã‡äz, 0 é¿ê—ã‡äz," & vbCrLf
        If CisFun.Mid2(GTitleDay(0), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
        Else
           gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(1), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
        Else
           gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(2), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
        Else
           gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(3), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
        Else
           gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(4), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
        Else
           gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(5), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "" & vbCrLf
        Else
           gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "" & vbCrLf
        End If
        If UKB_IRN = 0 Then
           gSL_Select = gSL_Select & "   FROM îÑè„åüé˚èÛãµä«óùÉrÉÖÅ[" & vbCrLf
        Else
           gSL_Select = gSL_Select & "   FROM îÑè„èoâ◊èÛãµä«óùÉrÉÖÅ[" & vbCrLf
        End If
        gSL_Select = gSL_Select & wWhereSS & vbCrLf
        gSL_Select = gSL_Select & "  GROUP BY å_ñÒêÊ, å_ñÒêÊñº, ï“èWå_ñÒêÊ" & vbCrLf
        gSL_Select = gSL_Select & " UNION" & vbCrLf
    End If
    
    gSL_Select = gSL_Select & " SELECT å_ñÒêÊ, å_ñÒêÊñº, ï“èWå_ñÒêÊ, SUM(ó\ë™ã‡äz) ó\ë™ã‡äz, SUM(é¿ê—ã‡äz) é¿ê—ã‡äz," & vbCrLf
    gSL_Select = gSL_Select & "        0, 0, 0, 0, 0, 0" & vbCrLf
    gSL_Select = gSL_Select & "   FROM îÑè„ì‡é¶èÛãµä«óùÉrÉÖÅ[" & vbCrLf
    gSL_Select = gSL_Select & wWhereN & vbCrLf
    gSL_Select = gSL_Select & "  GROUP BY å_ñÒêÊ, å_ñÒêÊñº, ï“èWå_ñÒêÊ) SY" & vbCrLf
    
    gSL_Select = gSL_Select & " GROUP BY å_ñÒêÊ, å_ñÒêÊñº, ï“èWå_ñÒêÊ"
    gSL_Select = gSL_Select & " ORDER BY ï“èWå_ñÒêÊ" & vbCrLf
    
    CisFun.ClipPut gSL_Select, False
    
    If Not USVRead(gSL_Select) Then Exit Function
    
    CXU0520_SQL1 = gSL_Select
    
    If Not BodySet Then Exit Function
    
'====================================== 2009/02/10 Start
    gSL_Select = " SELECT å_ñÒêÊ, å_ñÒêÊñº, ï“èWå_ñÒêÊ, SUM(ó\ë™ã‡äz) ó\ë™ã‡äz, SUM(é¿ê—ã‡äz) é¿ê—ã‡äz," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & ") é¿ê—ã‡äz0" & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & ") é¿ê—ã‡äz1" & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & ") é¿ê—ã‡äz2" & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & ") é¿ê—ã‡äz3" & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & ") é¿ê—ã‡äz4" & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & ") é¿ê—ã‡äz5" & "," & vbCrLf
    gSL_Select = gSL_Select & "        CASE WHEN SUM(ó\ë™ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
    gSL_Select = gSL_Select & "             CASE WHEN SUM(é¿ê—ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
    gSL_Select = gSL_Select & "                  ROUND(SUM(é¿ê—ã‡äz) / SUM(ó\ë™ã‡äz) * 100, 1)" & vbCrLf
    gSL_Select = gSL_Select & "        END  END êiíª," & vbCrLf
    If wKadou = 0 Then
       gSL_Select = gSL_Select & "        0 ì˙ìñÇËêî"
    Else
       gSL_Select = gSL_Select & "        ROUND(SUM(ó\ë™ã‡äz) / " & wKadou & ",1) ì˙ìñÇËêî"
    End If
    gSL_Select = gSL_Select & "   FROM ("
        
    gSL_Select = gSL_Select & " SELECT å_ñÒêÊ, å_ñÒêÊñº, ï“èWå_ñÒêÊ, 0 ó\ë™ã‡äz, SUM(é¿ê—ã‡äz) é¿ê—ã‡äz," & vbCrLf
    If CisFun.Mid2(GTitleDay(0), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
    Else
       gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(1), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
    Else
       gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(2), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
    Else
       gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(3), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
    Else
       gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(4), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
    Else
       gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(5), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "" & vbCrLf
    Else
       gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "" & vbCrLf
    End If
    If UKB_IRN = 0 Then
       gSL_Select = gSL_Select & "   FROM îÑè„åüé˚èÛãµä«óùÉrÉÖÅ[" & vbCrLf
    Else
       gSL_Select = gSL_Select & "   FROM îÑè„èoâ◊èÛãµä«óùÉrÉÖÅ[" & vbCrLf
    End If
    gSL_Select = gSL_Select & wWhereS & vbCrLf
    gSL_Select = gSL_Select & "  GROUP BY å_ñÒêÊ, å_ñÒêÊñº, ï“èWå_ñÒêÊ" & vbCrLf
    gSL_Select = gSL_Select & " UNION" & vbCrLf
    
    If CisFun.Val2(CisFun.Mid2(iH1_YMD, 7, 2)) < 7 Or CisFun.Val2(CisFun.Mid2(iH1_YMD, 7, 2)) > 25 Then
        gSL_Select = gSL_Select & " SELECT å_ñÒêÊ, å_ñÒêÊñº, ï“èWå_ñÒêÊ, 0 ó\ë™ã‡äz, 0 é¿ê—ã‡äz," & vbCrLf
        If CisFun.Mid2(GTitleDay(0), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
        Else
           gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(1), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
        Else
           gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(2), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
        Else
           gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(3), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
        Else
           gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(4), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
        Else
           gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(5), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "" & vbCrLf
        Else
           gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "" & vbCrLf
        End If
        If UKB_IRN = 0 Then
           gSL_Select = gSL_Select & "   FROM îÑè„åüé˚èÛãµä«óùÉrÉÖÅ[" & vbCrLf
        Else
           gSL_Select = gSL_Select & "   FROM îÑè„èoâ◊èÛãµä«óùÉrÉÖÅ[" & vbCrLf
        End If
        gSL_Select = gSL_Select & wWhereSS & vbCrLf
        gSL_Select = gSL_Select & "  GROUP BY å_ñÒêÊ, å_ñÒêÊñº, ï“èWå_ñÒêÊ" & vbCrLf
        gSL_Select = gSL_Select & " UNION" & vbCrLf
    End If
    
    gSL_Select = gSL_Select & " SELECT å_ñÒêÊ, å_ñÒêÊñº, ï“èWå_ñÒêÊ, SUM(ó\ë™ã‡äz) ó\ë™ã‡äz, SUM(é¿ê—ã‡äz) é¿ê—ã‡äz," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "" & vbCrLf
    gSL_Select = gSL_Select & "   FROM îÑè„ì‡é¶èÛãµä«óùÉrÉÖÅ[" & vbCrLf
    gSL_Select = gSL_Select & wWhereN & vbCrLf
    gSL_Select = gSL_Select & "  GROUP BY å_ñÒêÊ, å_ñÒêÊñº, ï“èWå_ñÒêÊ) SY" & vbCrLf
    
    gSL_Select = gSL_Select & " GROUP BY å_ñÒêÊ, å_ñÒêÊñº, ï“èWå_ñÒêÊ"
    gSL_Select = gSL_Select & " ORDER BY ï“èWå_ñÒêÊ" & vbCrLf
    
    CXU0520_SQL3 = gSL_Select
'====================================== 2009/02/10 End
    
    DBInput = True
    
End Function


'+--------------------------+
'+     ÉfÅ[É^ÉxÅ[ÉXï\é¶2    +
'+--------------------------+
Private Function DBInput2() As Boolean
    Dim wKadou      As Integer
    Dim wCnt        As Integer
    Dim wWhereS     As String
    Dim wWhereN     As String
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/04/01 INSERT START
    Dim wWhereSS    As String
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/04/01 INSERT END
    DBInput2 = False
    
    Call TitleDsp
    
    Sl_Where = "": gSL_Select = ""
' ï\é¶èåè ÉZÉbÉg
    With CisFun
        '+-------- èoâ◊ & ì‡é¶ÉfÅ[É^ --------+
        .Where_CreateKB = NewDocument
        ' éwíËì˙
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/12 UPDATE START
        '.Where_Create AlphaMe, "KS.é¿ê—îNåé", Mid(iH1_YMD, 1, 6)
        .Where_Create AlphaMe, "îNåé", Mid(iH1_YMD, 1, 6)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/12 UPDATE END
        ' å_ñÒêÊ
        .Where_Create AlphaMe, "å_ñÒêÊ", iH1_Torcd
         ' é©é–çHãÊ
         Call CisComb1.CodeGet
         .Where_Create AlphaMe, "é©é–çHãÊ", CisComb1.Code
         
         wWhereS = CisFun.Where_Phrase
         
    End With

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/12 INSERT START
    '+----------- ì‡é¶ÉfÅ[É^ÇÃÇ› -------------+
    With CisFun
        .Where_CreateKB = NewDocument
         ' îNåé
        .Where_Create AlphaMe, "îNåé", CisFun.Mid2(iH1_YMD, 1, 6)
         ' å_ñÒêÊ
         .Where_Create AlphaMe, "å_ñÒêÊ", iH1_Torcd
         ' é©é–çHãÊ
         Call CisComb1.CodeGet
         .Where_Create AlphaMe, "é©é–çHãÊ", CisComb1.Code
'         ' èoâ◊Ç»Çµ
'         .Where_Create Numeric, "èoâ◊éwé¶èëNO", 0, , Create
         
         wWhereN = CisFun.Where_Phrase
    End With
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/04/01 INSERT START
' ï\é¶èåè ÉZÉbÉg
    If CisFun.Val2(CisFun.Mid2(iH1_YMD, 7, 2)) < 7 Or CisFun.Val2(CisFun.Mid2(iH1_YMD, 7, 2)) > 25 Then
        
        CisFun.DateE_BefVal = iH1_YMD
        CisFun.DateEdit (Slash)
        If CisFun.Val2(CisFun.Mid2(iH1_YMD, 7, 2)) < 7 Then
           Call CisFun.Date_Add(iH1_YMD, M, -1)
        Else
           Call CisFun.Date_Add(iH1_YMD, D, 1)
        End If
        
        With CisFun
            '+-------- èoâ◊ & ì‡é¶ÉfÅ[É^ --------+
            .Where_CreateKB = NewDocument
    
            ' éwíËì˙
            .Where_Create AlphaMe, "îNåé", Mid(Replace(.Date_Result, "/", ""), 1, 6)
            ' å_ñÒêÊ
            .Where_Create AlphaMe, "å_ñÒêÊ", iH1_Torcd
             ' é©é–çHãÊ
             Call CisComb1.CodeGet
             .Where_Create AlphaMe, "é©é–çHãÊ", CisComb1.Code
             
             wWhereSS = CisFun.Where_Phrase
             
        End With
    End If
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/04/01 INSERT END

    
    '+------- â“ì≠ì˙êîéZèo ------+
    gSL_Select = ""
    gSL_Select = gSL_Select & " SELECT â“ìÆì˙êî FROM ÉJÉåÉìÉ_É}ÉXÉ^"
    gSL_Select = gSL_Select & "  WHERE éÊà¯êÊ = ''"
    gSL_Select = gSL_Select & "    AND éÊà¯êÊãÊï™ = 0"
    gSL_Select = gSL_Select & "    AND îN = '" & CisFun.Mid2(iH1_YMD, 1, 4) & "'"
    gSL_Select = gSL_Select & "    AND åé = '" & CisFun.Mid2(iH1_YMD, 5, 2) & "'"
    If CLMRead(gSL_Select) Then
       wKadou = CLM.â“ìÆì˙êî
    Else
       wKadou = 0
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/12 INSERT END
    
    gSL_Select = " SELECT é©é–çHãÊ, é©é–çHãÊñº, SUM(ó\ë™ã‡äz) ó\ë™ã‡äz, SUM(é¿ê—ã‡äz) é¿ê—ã‡äz," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
    gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "," & vbCrLf
    gSL_Select = gSL_Select & "        CASE WHEN SUM(ó\ë™ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
    gSL_Select = gSL_Select & "             CASE WHEN SUM(é¿ê—ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
    gSL_Select = gSL_Select & "                  ROUND(SUM(é¿ê—ã‡äz) / SUM(ó\ë™ã‡äz) * 100, 1)" & vbCrLf
    gSL_Select = gSL_Select & "        END  END êiíª," & vbCrLf
    If wKadou = 0 Then
       gSL_Select = gSL_Select & "        0 ì˙ìñÇËêî" & vbCrLf
    Else
       gSL_Select = gSL_Select & "        ROUND(SUM(ó\ë™ã‡äz) / " & wKadou & ",1) ì˙ìñÇËêî" & vbCrLf
    End If
    gSL_Select = gSL_Select & "   FROM (" & vbCrLf
    
    gSL_Select = gSL_Select & " SELECT é©é–çHãÊ, ISNULL(SY.ílñºèÃ, '') é©é–çHãÊñº, 0 ó\ë™ã‡äz, SUM(é¿ê—ã‡äz) é¿ê—ã‡äz," & vbCrLf
    If CisFun.Mid2(GTitleDay(0), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
    Else
       gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(1), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
    Else
       gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(2), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
    Else
       gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(3), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
    Else
       gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(4), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
    Else
       gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(5), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "," & vbCrLf
    Else
       gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "," & vbCrLf
    End If
    gSL_Select = gSL_Select & "        CASE WHEN SUM(ó\ë™ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
    gSL_Select = gSL_Select & "             CASE WHEN SUM(é¿ê—ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
    gSL_Select = gSL_Select & "                  ROUND(SUM(é¿ê—ã‡äz) / SUM(ó\ë™ã‡äz) * 100, 1)" & vbCrLf
    gSL_Select = gSL_Select & "        END  END êiíª," & vbCrLf
    If wKadou = 0 Then
       gSL_Select = gSL_Select & "        0 ì˙ìñÇËêî" & vbCrLf
    Else
       gSL_Select = gSL_Select & "        ROUND(SUM(ó\ë™ã‡äz) / " & wKadou & ",1) ì˙ìñÇËêî"
    End If
    If UKB_IRN = 0 Then
       gSL_Select = gSL_Select & "   FROM îÑè„åüé˚èÛãµä«óùÉrÉÖÅ[" & vbCrLf
    Else
       gSL_Select = gSL_Select & "   FROM îÑè„èoâ◊èÛãµä«óùÉrÉÖÅ[" & vbCrLf
    End If
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN ñºèÃÉ}ÉXÉ^ SY" & vbCrLf
    gSL_Select = gSL_Select & "             ON SY.ãÊï™ñºèÃ = 'é©é–çHãÊ'" & vbCrLf
    gSL_Select = gSL_Select & "            AND SY.íl = é©é–çHãÊ" & vbCrLf
    gSL_Select = gSL_Select & "            AND SY.ãÊï™É^ÉCÉv = ''" & vbCrLf
    gSL_Select = gSL_Select & wWhereS & vbCrLf
    gSL_Select = gSL_Select & "  GROUP BY é©é–çHãÊ, SY.ílñºèÃ" & vbCrLf
    
    If CisFun.Val2(CisFun.Mid2(iH1_YMD, 7, 2)) < 7 Or CisFun.Val2(CisFun.Mid2(iH1_YMD, 7, 2)) > 25 Then
        gSL_Select = gSL_Select & " UNION" & vbCrLf
        
        gSL_Select = gSL_Select & " SELECT é©é–çHãÊ, ISNULL(SY.ílñºèÃ, '') é©é–çHãÊñº, 0 ó\ë™ã‡äz, 0 é¿ê—ã‡äz," & vbCrLf
        If CisFun.Mid2(GTitleDay(0), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
        Else
           gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(1), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
        Else
           gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(2), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
        Else
           gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(3), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
        Else
           gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(4), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
        Else
           gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(5), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           gSL_Select = gSL_Select & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "," & vbCrLf
        Else
           gSL_Select = gSL_Select & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "," & vbCrLf
        End If
        gSL_Select = gSL_Select & "        CASE WHEN SUM(ó\ë™ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
        gSL_Select = gSL_Select & "             CASE WHEN SUM(é¿ê—ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
        gSL_Select = gSL_Select & "                  ROUND(SUM(é¿ê—ã‡äz) / SUM(ó\ë™ã‡äz) * 100, 1)" & vbCrLf
        gSL_Select = gSL_Select & "        END  END êiíª," & vbCrLf
        If wKadou = 0 Then
           gSL_Select = gSL_Select & "        0 ì˙ìñÇËêî" & vbCrLf
        Else
           gSL_Select = gSL_Select & "        ROUND(SUM(ó\ë™ã‡äz) / " & wKadou & ",1) ì˙ìñÇËêî"
        End If
        If UKB_IRN = 0 Then
           gSL_Select = gSL_Select & "   FROM îÑè„åüé˚èÛãµä«óùÉrÉÖÅ[" & vbCrLf
        Else
           gSL_Select = gSL_Select & "   FROM îÑè„èoâ◊èÛãµä«óùÉrÉÖÅ[" & vbCrLf
        End If
        gSL_Select = gSL_Select & "        LEFT OUTER JOIN ñºèÃÉ}ÉXÉ^ SY" & vbCrLf
        gSL_Select = gSL_Select & "             ON SY.ãÊï™ñºèÃ = 'é©é–çHãÊ'" & vbCrLf
        gSL_Select = gSL_Select & "            AND SY.íl = é©é–çHãÊ" & vbCrLf
        gSL_Select = gSL_Select & "            AND SY.ãÊï™É^ÉCÉv = ''" & vbCrLf
        gSL_Select = gSL_Select & wWhereSS & vbCrLf
        gSL_Select = gSL_Select & "  GROUP BY é©é–çHãÊ, SY.ílñºèÃ" & vbCrLf
    End If
    gSL_Select = gSL_Select & " UNION" & vbCrLf
    
    gSL_Select = gSL_Select & " SELECT é©é–çHãÊ, ISNULL(SY.ílñºèÃ, '') é©é–çHãÊñº, SUM(ó\ë™ã‡äz) ó\ë™ã‡äz, SUM(é¿ê—ã‡äz) é¿ê—ã‡äz," & vbCrLf
    gSL_Select = gSL_Select & "        0, 0, 0, 0, 0, 0," & vbCrLf
    gSL_Select = gSL_Select & "        CASE WHEN SUM(ó\ë™ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
    gSL_Select = gSL_Select & "             CASE WHEN SUM(é¿ê—ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
    gSL_Select = gSL_Select & "                  ROUND(SUM(é¿ê—ã‡äz) / SUM(ó\ë™ã‡äz) * 100, 1)" & vbCrLf
    gSL_Select = gSL_Select & "        END  END êiíª," & vbCrLf
    If wKadou = 0 Then
       gSL_Select = gSL_Select & "        0 ì˙ìñÇËêî" & vbCrLf
    Else
       gSL_Select = gSL_Select & "        ROUND(SUM(ó\ë™ã‡äz) / " & wKadou & ",1) ì˙ìñÇËêî" & vbCrLf
    End If
    gSL_Select = gSL_Select & "   FROM îÑè„ì‡é¶èÛãµä«óùÉrÉÖÅ[" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN ñºèÃÉ}ÉXÉ^ SY" & vbCrLf
    gSL_Select = gSL_Select & "             ON SY.ãÊï™ñºèÃ = 'é©é–çHãÊ'" & vbCrLf
    gSL_Select = gSL_Select & "            AND SY.íl = é©é–çHãÊ" & vbCrLf
    gSL_Select = gSL_Select & "            AND SY.ãÊï™É^ÉCÉv = ''" & vbCrLf
    gSL_Select = gSL_Select & wWhereN & vbCrLf
    gSL_Select = gSL_Select & "  GROUP BY é©é–çHãÊ, SY.ílñºèÃ) SY" & vbCrLf
    
    gSL_Select = gSL_Select & "  GROUP BY é©é–çHãÊ, é©é–çHãÊñº" & vbCrLf
    gSL_Select = gSL_Select & " ORDER BY é©é–çHãÊ" & vbCrLf
    
    CisFun.ClipPut gSL_Select, False
    
    If Not USVRead(gSL_Select) Then Exit Function
        
    If Not BodySet2 Then Exit Function
    
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/02/12 INSERT START
    CXU0520_SQL4 = " SELECT é©é–çHãÊ, é©é–çHãÊñº, SUM(ó\ë™ã‡äz) ó\ë™ã‡äz, SUM(é¿ê—ã‡äz) é¿ê—ã‡äz," & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & ") é¿ê—ã‡äz0," & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & ") é¿ê—ã‡äz1," & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & ") é¿ê—ã‡äz2," & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & ") é¿ê—ã‡äz3," & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & ") é¿ê—ã‡äz4," & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & ") é¿ê—ã‡äz5," & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "        CASE WHEN SUM(ó\ë™ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "             CASE WHEN SUM(é¿ê—ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "                  ROUND(SUM(é¿ê—ã‡äz) / SUM(ó\ë™ã‡äz) * 100, 1)" & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "        END  END êiíª," & vbCrLf
    If wKadou = 0 Then
       CXU0520_SQL4 = CXU0520_SQL4 & "        0 ì˙ìñÇËêî"
    Else
       CXU0520_SQL4 = CXU0520_SQL4 & "        ROUND(SUM(ó\ë™ã‡äz) / " & wKadou & ",1) ì˙ìñÇËêî"
    End If
    CXU0520_SQL4 = CXU0520_SQL4 & "   FROM ("
    
    CXU0520_SQL4 = CXU0520_SQL4 & " SELECT é©é–çHãÊ, ISNULL(SY.ílñºèÃ, '') é©é–çHãÊñº, 0 ó\ë™ã‡äz, SUM(é¿ê—ã‡äz) é¿ê—ã‡äz," & vbCrLf
    If CisFun.Mid2(GTitleDay(0), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       CXU0520_SQL4 = CXU0520_SQL4 & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
    Else
       CXU0520_SQL4 = CXU0520_SQL4 & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(1), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       CXU0520_SQL4 = CXU0520_SQL4 & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
    Else
       CXU0520_SQL4 = CXU0520_SQL4 & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(2), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       CXU0520_SQL4 = CXU0520_SQL4 & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
    Else
       CXU0520_SQL4 = CXU0520_SQL4 & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(3), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       CXU0520_SQL4 = CXU0520_SQL4 & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
    Else
       CXU0520_SQL4 = CXU0520_SQL4 & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(4), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       CXU0520_SQL4 = CXU0520_SQL4 & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
    Else
       CXU0520_SQL4 = CXU0520_SQL4 & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
    End If
    If CisFun.Mid2(GTitleDay(5), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
       CXU0520_SQL4 = CXU0520_SQL4 & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "," & vbCrLf
    Else
       CXU0520_SQL4 = CXU0520_SQL4 & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "," & vbCrLf
    End If
    CXU0520_SQL4 = CXU0520_SQL4 & "        CASE WHEN SUM(ó\ë™ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "             CASE WHEN SUM(é¿ê—ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "                  ROUND(SUM(é¿ê—ã‡äz) / SUM(ó\ë™ã‡äz) * 100, 1)" & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "        END  END êiíª," & vbCrLf
    If wKadou = 0 Then
       CXU0520_SQL4 = CXU0520_SQL4 & "        0 ì˙ìñÇËêî"
    Else
       CXU0520_SQL4 = CXU0520_SQL4 & "        ROUND(SUM(ó\ë™ã‡äz) / " & wKadou & ",1) ì˙ìñÇËêî"
    End If
    If UKB_IRN = 0 Then
       CXU0520_SQL4 = CXU0520_SQL4 & "   FROM îÑè„åüé˚èÛãµä«óùÉrÉÖÅ[" & vbCrLf
    Else
       CXU0520_SQL4 = CXU0520_SQL4 & "   FROM îÑè„èoâ◊èÛãµä«óùÉrÉÖÅ[" & vbCrLf
    End If
    CXU0520_SQL4 = CXU0520_SQL4 & "        LEFT OUTER JOIN ñºèÃÉ}ÉXÉ^ SY"
    CXU0520_SQL4 = CXU0520_SQL4 & "             ON SY.ãÊï™ñºèÃ = 'é©é–çHãÊ'"
    CXU0520_SQL4 = CXU0520_SQL4 & "            AND SY.íl = é©é–çHãÊ"
    CXU0520_SQL4 = CXU0520_SQL4 & "            AND SY.ãÊï™É^ÉCÉv = ''"
    CXU0520_SQL4 = CXU0520_SQL4 & wWhereS & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "  GROUP BY é©é–çHãÊ, SY.ílñºèÃ" & vbCrLf
    
    CXU0520_SQL4 = CXU0520_SQL4 & " UNION" & vbCrLf
    
    If CisFun.Val2(CisFun.Mid2(iH1_YMD, 7, 2)) < 7 Or CisFun.Val2(CisFun.Mid2(iH1_YMD, 7, 2)) > 25 Then
        CXU0520_SQL4 = CXU0520_SQL4 & " SELECT é©é–çHãÊ, ISNULL(SY.ílñºèÃ, '') é©é–çHãÊñº, 0 ó\ë™ã‡äz, 0 é¿ê—ã‡äz," & vbCrLf
        If CisFun.Mid2(GTitleDay(0), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           CXU0520_SQL4 = CXU0520_SQL4 & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
        Else
           CXU0520_SQL4 = CXU0520_SQL4 & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(0), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(1), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           CXU0520_SQL4 = CXU0520_SQL4 & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
        Else
           CXU0520_SQL4 = CXU0520_SQL4 & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(1), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(2), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           CXU0520_SQL4 = CXU0520_SQL4 & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
        Else
           CXU0520_SQL4 = CXU0520_SQL4 & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(2), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(3), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           CXU0520_SQL4 = CXU0520_SQL4 & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
        Else
           CXU0520_SQL4 = CXU0520_SQL4 & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(3), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(4), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           CXU0520_SQL4 = CXU0520_SQL4 & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
        Else
           CXU0520_SQL4 = CXU0520_SQL4 & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(4), 4, 2)) & "," & vbCrLf
        End If
        If CisFun.Mid2(GTitleDay(5), 1, 2) = CisFun.Mid2(iH1_YMD, 5, 2) Then
           CXU0520_SQL4 = CXU0520_SQL4 & "        0 é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "," & vbCrLf
        Else
           CXU0520_SQL4 = CXU0520_SQL4 & "        SUM(é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & ") é¿ê—ã‡äz" & CisFun.Val2(CisFun.Mid2(GTitleDay(5), 4, 2)) & "," & vbCrLf
        End If
        CXU0520_SQL4 = CXU0520_SQL4 & "        CASE WHEN SUM(ó\ë™ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
        CXU0520_SQL4 = CXU0520_SQL4 & "             CASE WHEN SUM(é¿ê—ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
        CXU0520_SQL4 = CXU0520_SQL4 & "                  ROUND(SUM(é¿ê—ã‡äz) / SUM(ó\ë™ã‡äz) * 100, 1)" & vbCrLf
        CXU0520_SQL4 = CXU0520_SQL4 & "        END  END êiíª," & vbCrLf
        If wKadou = 0 Then
           CXU0520_SQL4 = CXU0520_SQL4 & "        0 ì˙ìñÇËêî"
        Else
           CXU0520_SQL4 = CXU0520_SQL4 & "        ROUND(SUM(ó\ë™ã‡äz) / " & wKadou & ",1) ì˙ìñÇËêî"
        End If
        If UKB_IRN = 0 Then
           CXU0520_SQL4 = CXU0520_SQL4 & "   FROM îÑè„åüé˚èÛãµä«óùÉrÉÖÅ[" & vbCrLf
        Else
           CXU0520_SQL4 = CXU0520_SQL4 & "   FROM îÑè„èoâ◊èÛãµä«óùÉrÉÖÅ[" & vbCrLf
        End If
        CXU0520_SQL4 = CXU0520_SQL4 & "        LEFT OUTER JOIN ñºèÃÉ}ÉXÉ^ SY"
        CXU0520_SQL4 = CXU0520_SQL4 & "             ON SY.ãÊï™ñºèÃ = 'é©é–çHãÊ'"
        CXU0520_SQL4 = CXU0520_SQL4 & "            AND SY.íl = é©é–çHãÊ"
        CXU0520_SQL4 = CXU0520_SQL4 & "            AND SY.ãÊï™É^ÉCÉv = ''"
        CXU0520_SQL4 = CXU0520_SQL4 & wWhereSS & vbCrLf
        CXU0520_SQL4 = CXU0520_SQL4 & "  GROUP BY é©é–çHãÊ, SY.ílñºèÃ" & vbCrLf
        
        CXU0520_SQL4 = CXU0520_SQL4 & " UNION" & vbCrLf
    End If
    CXU0520_SQL4 = CXU0520_SQL4 & " SELECT é©é–çHãÊ, ISNULL(SY.ílñºèÃ, '') é©é–çHãÊñº, SUM(ó\ë™ã‡äz) ó\ë™ã‡äz, SUM(é¿ê—ã‡äz) é¿ê—ã‡äz," & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "        0,0,0,0,0,0," & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "        CASE WHEN SUM(ó\ë™ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "             CASE WHEN SUM(é¿ê—ã‡äz) = 0 THEN 0 ELSE" & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "                  ROUND(SUM(é¿ê—ã‡äz) / SUM(ó\ë™ã‡äz) * 100, 1)" & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "        END  END êiíª," & vbCrLf
    If wKadou = 0 Then
       CXU0520_SQL4 = CXU0520_SQL4 & "        0 ì˙ìñÇËêî"
    Else
       CXU0520_SQL4 = CXU0520_SQL4 & "        ROUND(SUM(ó\ë™ã‡äz) / " & wKadou & ",1) ì˙ìñÇËêî"
    End If
    CXU0520_SQL4 = CXU0520_SQL4 & "   FROM îÑè„ì‡é¶èÛãµä«óùÉrÉÖÅ[" & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "        LEFT OUTER JOIN ñºèÃÉ}ÉXÉ^ SY"
    CXU0520_SQL4 = CXU0520_SQL4 & "             ON SY.ãÊï™ñºèÃ = 'é©é–çHãÊ'"
    CXU0520_SQL4 = CXU0520_SQL4 & "            AND SY.íl = é©é–çHãÊ"
    CXU0520_SQL4 = CXU0520_SQL4 & "            AND SY.ãÊï™É^ÉCÉv = ''"
    CXU0520_SQL4 = CXU0520_SQL4 & wWhereN & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & "  GROUP BY é©é–çHãÊ, SY.ílñºèÃ) SY" & vbCrLf
    
    CXU0520_SQL4 = CXU0520_SQL4 & "  GROUP BY é©é–çHãÊ, é©é–çHãÊñº" & vbCrLf
    CXU0520_SQL4 = CXU0520_SQL4 & " ORDER BY é©é–çHãÊ" & vbCrLf
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/02/12 INSERT END
    
    DBInput2 = True
End Function
'******************************************
'*     ñºèÃÉ}ÉXÉ^ÇÊÇËçHèÍíäèo Åï ï\é¶     *
'******************************************
Private Sub KojoDsp()
    Dim wCnt1   As Integer
    Dim wRow    As Integer
    
    With VsGrid2
       .Row = 0
       .TextMatrix(.Row, 0) = "Å@á@çáÅ@Å@åv"
       For wCnt1 = 0 To 10
           .Col = wCnt1
           .CellBackColor = G2Color(0).BackColor
       Next wCnt1
    
       wRow = 1
       For wCnt1 = 1 To 10
           .TextMatrix(0, wCnt1) = ""
       Next wCnt1
       
       gSL_Select = ""
       gSL_Select = gSL_Select & " SELECT * FROM ñºèÃÉ}ÉXÉ^"
       gSL_Select = gSL_Select & "  WHERE ãÊï™ñºèÃ = 'é©é–çHãÊ'"
       gSL_Select = gSL_Select & "    AND ãÊï™É^ÉCÉv = ''"
       If SYMRead(gSL_Select) Then
          If .Rows - 1 < CisDB.RecordCount Then
             .Rows = CisDB.RecordCount + 1
          End If
          Do Until Not SYM_RDSTS
             .TextMatrix(wRow, 0) = "Å@Å@" & SYM.ílñºèÃ
             For wCnt1 = 0 To 10
                 .Col = wCnt1
                 .Row = wRow
                 .CellBackColor = G2Color(wRow).BackColor
                 If wCnt1 <> 0 Then
                    .TextMatrix(.Row, .Col) = ""
                 End If
             Next wCnt1
             Call SYMReadNext
             wRow = wRow + 1
          Loop
       End If
    
    End With
End Sub

'********************************
'*     ÉOÉäÉbÉhÉ^ÉCÉgÉãï\é¶     *
'********************************
Private Sub TitleDsp()
    Dim wWeekCnt        As Integer
    Dim wCntDay         As Integer
    Dim wGridCol        As Integer
    
    GTitle(1) = Mid(iH1_YMD, 5, 2) & "åéìxó\ë™"
    GTitle(2) = Mid(iH1_YMD, 5, 2) & "åéìxé¿ê—"
    GTitle(3) = "êiíª"
    
    With CisFun
         .DateE_BefVal = iH1_YMD
         .DateEdit (Slash)
         wWeekCnt = 2 - Weekday(CDate(.DateE_AfVal))
         
         wGridCol = 0
         For wCntDay = wWeekCnt To wWeekCnt + 5
            .Date_RsFormat = fDate
            If wCntDay = 0 Then
               GTitleDay(wGridCol) = Mid(iH1_YMD, 5, 2) & "/" & Mid(iH1_YMD, 7, 2)
            Else
               Call .Date_Add(iH1_YMD, D, wCntDay)
               GTitleDay(wGridCol) = Mid(.Date_Result, 6, 5)
            End If
            wGridCol = wGridCol + 1
         Next wCntDay
    End With
End Sub
