VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Begin VB.Form COK0010 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'ŒÅ’è(Àü)
   Caption         =   "x•¥“ü—Í"
   ClientHeight    =   11025
   ClientLeft      =   -1110
   ClientTop       =   1545
   ClientWidth     =   15270
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "‚l‚r ƒSƒVƒbƒN"
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
   PaletteMode     =   1  'Z µ°ÀŞ°
   ScaleHeight     =   11025
   ScaleWidth      =   15270
   Begin VB.PictureBox B3_Area2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   6570
      ScaleHeight     =   1455
      ScaleWidth      =   8220
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1605
      Width           =   8280
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   495
         Index           =   0
         Left            =   6045
         Top             =   435
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   873
         ForeColor       =   16711680
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin CisYMD_With_Btn.CisYMDwB iB3_SYmd 
            Height          =   345
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   75
            Width           =   1905
            _ExtentX        =   2831
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Object.Height          =   345
            Object.Width           =   1905
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   495
         Index           =   0
         Left            =   495
         Top             =   435
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   873
         ForeColor       =   16711680
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin Cis3D_v60.CIS3D B3lb_Paynm 
            Height          =   330
            Index           =   1
            Left            =   510
            Top             =   60
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin CisText_V60.CisText iB3_Paykb 
            Height          =   360
            Index           =   1
            Left            =   150
            TabIndex        =   8
            Top             =   60
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   495
         Index           =   0
         Left            =   3240
         Top             =   435
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   873
         ForeColor       =   16711680
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin CisText_V60.CisText iB3_SKin 
            Height          =   360
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   60
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   635
            cFormat         =   "###,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "###,###,###"
            MaxLength       =   11
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "#########"
            cILength        =   9
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D23 
         Height          =   495
         Index           =   0
         Left            =   5025
         Top             =   435
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   873
         ForeColor       =   16711680
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin CisText_V60.CisText iB3_Tesu 
            Height          =   360
            Index           =   1
            Left            =   255
            TabIndex        =   11
            Top             =   60
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   635
            cFormat         =   "###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
            cFdAutoFormat   =   1
            cGFormat        =   "###"
            cILength        =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D24 
         Height          =   495
         Index           =   0
         Left            =   2385
         Top             =   435
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         ForeColor       =   0
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cAlingnment     =   4
         cPositionX      =   -50
         Begin CisText_V60.CisText iB3_Ritu 
            Height          =   360
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   45
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   635
            cFormat         =   "###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
            cFdAutoFormat   =   1
            cGFormat        =   "###"
            cILength        =   3
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   390
         Index           =   0
         Left            =   2385
         Top             =   45
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   688
         ForeColor       =   16711680
         Caption         =   "—¦"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   390
         Index           =   0
         Left            =   495
         Top             =   45
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   688
         ForeColor       =   16711680
         Caption         =   "x•¥•û–@"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   390
         Index           =   1
         Left            =   6045
         Top             =   45
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   688
         ForeColor       =   16711680
         Caption         =   "x •¥ “ú"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   390
         Index           =   1
         Left            =   60
         Top             =   45
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   688
         ForeColor       =   0
         Caption         =   "‡‚"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   390
         Index           =   2
         Left            =   5025
         Top             =   45
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   688
         ForeColor       =   16711680
         Caption         =   "è”—¿"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   390
         Index           =   2
         Left            =   3240
         Top             =   45
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   688
         ForeColor       =   16711680
         Caption         =   "x•¥‹àŠz"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   495
         Index           =   3
         Left            =   60
         Top             =   435
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   873
         ForeColor       =   0
         Caption         =   "1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Height          =   495
         Index           =   1
         Left            =   6045
         Top             =   930
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   873
         ForeColor       =   16711680
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin CisYMD_With_Btn.CisYMDwB iB3_SYmd 
            Height          =   345
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   75
            Width           =   1905
            _ExtentX        =   2831
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Object.Height          =   345
            Object.Width           =   1905
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   495
         Index           =   1
         Left            =   495
         Top             =   930
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   873
         ForeColor       =   16711680
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin Cis3D_v60.CIS3D B3lb_Paynm 
            Height          =   330
            Index           =   2
            Left            =   510
            Top             =   60
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin CisText_V60.CisText iB3_Paykb 
            Height          =   360
            Index           =   2
            Left            =   150
            TabIndex        =   13
            Top             =   60
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   495
         Index           =   1
         Left            =   3240
         Top             =   930
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   873
         ForeColor       =   16711680
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin CisText_V60.CisText iB3_SKin 
            Height          =   360
            Index           =   2
            Left            =   120
            TabIndex        =   15
            Top             =   60
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   635
            cFormat         =   "###,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "###,###,###"
            MaxLength       =   11
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "#########"
            cILength        =   9
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D23 
         Height          =   495
         Index           =   1
         Left            =   5025
         Top             =   930
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   873
         ForeColor       =   16711680
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin CisText_V60.CisText iB3_Tesu 
            Height          =   360
            Index           =   2
            Left            =   255
            TabIndex        =   16
            Top             =   60
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   635
            cFormat         =   "###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
            cFdAutoFormat   =   1
            cGFormat        =   "###"
            cILength        =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D24 
         Height          =   495
         Index           =   1
         Left            =   2385
         Top             =   930
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         ForeColor       =   0
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cAlingnment     =   4
         cPositionX      =   -50
         Begin CisText_V60.CisText iB3_Ritu 
            Height          =   360
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   45
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   635
            cFormat         =   "###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
            cFdAutoFormat   =   1
            cGFormat        =   "###"
            cILength        =   3
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   495
         Index           =   4
         Left            =   60
         Top             =   930
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   873
         ForeColor       =   0
         Caption         =   "2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   2415
      ScaleHeight     =   900
      ScaleWidth      =   3840
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1620
      Width           =   3900
      Begin Cis3D_v60.CIS3D HLB_Hinnm 
         Height          =   825
         Left            =   45
         Top             =   45
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1455
         ForeColor       =   16711680
         Caption         =   "Œ_ –ñ æ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin Cis3D_v60.CIS3D B1lb_Tornm 
            Height          =   330
            Left            =   1125
            Top             =   360
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin CisText_V60.CisText iB1_Torcd 
            Height          =   360
            Left            =   135
            TabIndex        =   4
            Top             =   360
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   405
      Index           =   0
      Left            =   450
      ScaleHeight     =   345
      ScaleWidth      =   2010
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2715
      Width           =   2070
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   345
         Index           =   1
         Left            =   15
         Top             =   0
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   609
         Caption         =   " •\¦Œ”"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin Cis3D_v60.CIS3D Dsp_Kensu 
            Height          =   270
            Left            =   1095
            Top             =   30
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   476
            BackColor       =   16777152
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -60
         End
      End
   End
   Begin VB.PictureBox B2_Area1 
      Height          =   6555
      Left            =   420
      ScaleHeight     =   6495
      ScaleWidth      =   14385
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4335
      Width           =   14445
      Begin vsFlexLib.vsFlexArray VsGrid1 
         Height          =   6345
         Left            =   75
         TabIndex        =   18
         Top             =   75
         Width           =   14235
         _Version        =   196608
         _ExtentX        =   25109
         _ExtentY        =   11192
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   $"COK0010.frx":0000
         Rows            =   18
         Cols            =   32
         BackColor       =   12648447
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   330
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
   Begin VB.PictureBox B3_Area1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   435
      ScaleHeight     =   930
      ScaleWidth      =   14370
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3180
      Width           =   14430
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   855
         Left            =   75
         Top             =   45
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1508
         ForeColor       =   0
         Caption         =   "‘OŒŒJ‰z"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin Cis3D_v60.CIS3D B3lb_ZKuri 
            Height          =   330
            Left            =   90
            Top             =   360
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   582
            BackColor       =   12648384
            Caption         =   "###,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
            cPositionX      =   -50
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   855
         Left            =   12795
         Top             =   45
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1508
         ForeColor       =   0
         Caption         =   "¿‹‘‡‚"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin Cis3D_v60.CIS3D B3lb_SeiNo 
            Height          =   330
            Left            =   90
            Top             =   360
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "XXXXXXXXXX"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   855
         Left            =   1755
         Top             =   45
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   1508
         ForeColor       =   16711680
         Caption         =   "d“ü‹àŠz@@@Á”ïÅ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin CisText_V60.CisText iB3_SRKin 
            Height          =   360
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   635
            cFormat         =   "###,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "###,###,###"
            MaxLength       =   11
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "#########"
            cILength        =   9
         End
         Begin Cis3D_v60.CIS3D B3lb_SRZei 
            Height          =   330
            Left            =   1620
            Top             =   360
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   582
            BackColor       =   12648384
            Caption         =   "#,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
            cPositionX      =   -50
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D12 
         Height          =   855
         Left            =   9405
         Top             =   45
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   1508
         ForeColor       =   16711680
         Caption         =   "x•¥‹àŠz"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin CisText_V60.CisText iB3_SHKin 
            Height          =   360
            Left            =   90
            TabIndex        =   7
            Top             =   360
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   635
            cFormat         =   "###,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "###,###,###"
            MaxLength       =   11
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "#########"
            cILength        =   9
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   855
         Left            =   7725
         Top             =   45
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1508
         ForeColor       =   0
         Caption         =   "“–Œ¿‹‹àŠz"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin Cis3D_v60.CIS3D B3lb_Sekyu 
            Height          =   330
            Left            =   90
            Top             =   360
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   582
            BackColor       =   12648384
            Caption         =   "###,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
            cPositionX      =   -50
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D17 
         Height          =   855
         Left            =   11115
         Top             =   45
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1508
         ForeColor       =   0
         Caption         =   "ŒJ‰z‹àŠz"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin Cis3D_v60.CIS3D B3lb_TKuri 
            Height          =   330
            Left            =   90
            Top             =   360
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   582
            BackColor       =   12648384
            Caption         =   "###,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
            cPositionX      =   -50
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   855
         Left            =   4740
         Top             =   45
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   1508
         ForeColor       =   16711680
         Caption         =   "‘ŠE‹àŠz@@@Á”ïÅ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin CisText_V60.CisText iB3_SSKin 
            Height          =   360
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   635
            cFormat         =   "###,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "###,###,###"
            MaxLength       =   11
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "#########"
            cILength        =   9
         End
         Begin Cis3D_v60.CIS3D B3lb_SSZei 
            Height          =   330
            Left            =   1620
            Top             =   360
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   582
            BackColor       =   12648384
            Caption         =   "#,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
            cPositionX      =   -50
         End
      End
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   450
      ScaleHeight     =   900
      ScaleWidth      =   1860
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1620
      Width           =   1920
      Begin Cis3D_v60.CIS3D HLB_UYm 
         Height          =   825
         Left            =   60
         Top             =   45
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1455
         ForeColor       =   16711680
         Caption         =   "x•¥”NŒ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         Begin CisYM_V60.CisYM iH1_SYm 
            Height          =   390
            Left            =   150
            TabIndex        =   2
            Top             =   360
            Width           =   1485
            _ExtentX        =   2328
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   7905
      ScaleHeight     =   660
      ScaleWidth      =   6885
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   735
      Width           =   6945
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   4305
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   3
         BtnCaption      =   "æÁ"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r –¾’©"
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
         Left            =   5160
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   2
         BtnCaption      =   "“ü—Í"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r –¾’©"
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
         Left            =   6015
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   1
         BtnCaption      =   "I—¹"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r –¾’©"
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
         Left            =   2595
         Top             =   30
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   11
         BtnCaption      =   "ŒŸõ"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r –¾’©"
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   1740
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   4
         BtnCaption      =   "íœ"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r –¾’©"
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
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnCaption      =   "’Ç‰Á"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r –¾’©"
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
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   8
         BtnCaption      =   "C³"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r –¾’©"
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
      Begin CisBtn_60.CisBtn PB_PRT 
         Height          =   585
         Left            =   3450
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   13
         BtnCaption      =   "ˆóü"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "‚l‚r –¾’©"
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
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   150
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   540
      Width           =   345
   End
   Begin VB.Timer Timer1 
      Left            =   570
      Top             =   540
   End
   Begin VB.Timer Timer2 
      Left            =   1080
      Top             =   540
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
      Caption         =   "y@x•¥“ü—Í  z"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
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
         BorderStyle     =   0  '‚È‚µ
         Height          =   345
         Index           =   1
         Left            =   10620
         ScaleHeight     =   345
         ScaleWidth      =   3675
         TabIndex        =   20
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
            Caption         =   "íœ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
            Caption         =   "’Ç‰Á"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
            Caption         =   "C³"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "‚l‚r ƒSƒVƒbƒN"
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
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
            Name            =   "‚l‚r ƒSƒVƒbƒN"
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
Attribute VB_Name = "COK0010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************************'
'**
'**       ƒtƒH[ƒ€–¼    :   x•¥“ü—Í
'**       ƒtƒH[ƒ€ID    :   COK0010
'**       ˆ—ŠT—v      :
'**
'**       ì  ¬  “ú    :   2004/12/07  By CIS SIMPR-A
'**       •Ï  X  “ú    :
'**
'********************************************************************************************'
    Dim CisFrm          As CISFormContorl
'
    Dim FormAct         As Boolean      ' Ì«°Ñ±¸Ã¨ÌŞ Ì×¸Ş
    Dim ProcHB          As String * 2   ' “ü—ÍˆæØ‘Ö

    Dim SyoriKB         As String       ' ˆ—‹æ•ª(V‹K¥C³¥íœ “™)
    Dim SyoriNM         As String       ' ˆ—–¼Ì(V‹K¥C³¥íœ “™)
    Dim CisVsGrid1      As New CisVsGrid3
    Dim RowColFlg       As Boolean
'
    Dim wKurikosi       As Currency     '>ŒJ‰zŠÇ—
    Dim wZRitu          As Currency     '>Á”ïÅ—¦
    Dim wZHasu          As Currency     '>’[”ˆ—
    Dim wShime          As Byte         '>’÷“ú
    Dim wKin            As Double
    Dim wIndex          As Byte
    Dim wZenYm          As String
    Dim wYmd            As String       ' x•¥“ú‚ÌZo—pÜ°¸
    Dim wDay            As Byte         ' ª
    Dim wAddSu          As Integer      ' ª
    Dim SaveRow         As Integer
'+----------------------------------------------------------+
'+      ƒtƒH[ƒ€    s    ‚k‚‚‚„ƒCƒxƒ“ƒg    t              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   # “ñ d ‹N “® –h ~ #
'   #------------------#
    If Not Dupli_Start Then End
'   #------------------#
'   # ‰ Šú “à —e æ “¾ #
'   #------------------#
    If Not IniGet Then End
'   #------------------#
'   #  € –Ú ‰ Šú ‰»   #
'   #------------------#
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
'
    If gTantoName = "" Then
        Head_Tanto.Visible = False
    Else
        Head_Tanto.Visible = True
        Head_Tanto = gTantoName
    End If
'
    Dummy.Left = -1000
    Dummy.Enabled = False
'
    Timer1.Interval = 1000   ' •\¦ŠÔŠu
    Timer1.Enabled = True    ' •\¦ŠJn
    Timer2.Interval = 500    ' ˆ—‹æ•ª“_–ÅŠÔŠu(Šî–{)
    Timer2.Enabled = False   ' ˆ—‹æ•ª“_–Å‹Ö~(Šî–{)
'
    '+--------------------------------+
    '+  ƒOƒŠƒbƒh‰Šúİ’è(Ş—¿Ë•i”Ô)
    '+--------------------------------+
    With CisVsGrid1
        Set .GridObj = VsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 350
        .PatanMax = 1
        .SelectCol = 30
        .InitGet = False
        .Init
    End With
    Call GridInit1
'
'   #------------------#
'   #  ˆ—‹æ•ª‰Šú•\¦ #
'   #------------------#
    SyoriKB = ""
    Call SyoriName(SyoriKB)
'   #------------------#
'   #    ‰Šú“ü—Íİ’è   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   # ‰æ–Ê’†‰›ˆÊ’u•\¦   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)
'
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      ƒtƒH[ƒ€    s‚`‚ƒ‚”‚‰‚–‚‚”‚… ƒCƒxƒ“ƒgt                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
'   >> ŒJ‰zŠÇ—
    If MeisyoGet("”ƒŠ|ŠÇ—", "1") Then
        wKurikosi = SYM.”š1
    Else
        With CisFun
            .MB_Lines = 4
            .MB_Title = "ƒGƒ‰[‰æ–Ê"
            .MB_Button = Error
            .MB_MSG(1) = "     ŒJ‰zŠÇ—@–¢“o˜^Ii–¼Ìƒ}ƒXƒ^j            "
            .MB_MSG(3) = "     ˆ—‚ğI—¹‚µ‚Ü‚·I      "
            .MBOX
        End With
        Call PB_END_Click
        Exit Sub
    End If
'   >> Á”ïÅ
    If MeisyoGet("Á”ïÅ", "1") Then
        wZRitu = SYM.”š1
        wZHasu = SYM.”š2
    Else
        With CisFun
            .MB_Lines = 4
            .MB_Title = "ƒGƒ‰[‰æ–Ê"
            .MB_Button = Error
            .MB_MSG(1) = "     Á”ïÅ@–¢“o˜^Ii–¼Ìƒ}ƒXƒ^j            "
            .MB_MSG(3) = "     ˆ—‚ğI—¹‚µ‚Ü‚·I      "
            .MBOX
        End With
        Call PB_END_Click
        Exit Sub
    End If
'ˆ—Œ
    If MeisyoGet("ˆ—“à—e", "POK0110") Then
        iH1_SYm = Left(SYM.‰p”š1, 6)
    End If

    FormAct = Not FormAct
End Sub
'+-------------------------------------------------------------+
'+      ƒtƒH[ƒ€    s ‚j‚…‚™‚c‚‚—‚ƒCƒxƒ“ƒg t                 +
'+-------------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+-------------------------------------------------------------+
'+      ƒtƒH[ƒ€s ƒtƒ@ƒ“ƒNƒVƒ‡ƒ“ƒL[—pƒAƒNƒZƒ‰ƒŒ[ƒ^ t         +
'+-------------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Key_Acc = False
'
    Select Case Key_Code
           Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click    'yV‹Kz
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    'yC³z
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    'yíœz
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  'yŒŸõz
           Case vbKeyF8:       If PB_PRT.Visible Then Call PB_PRT_Click    'yˆóüz
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    'yæÁz
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    'yI—¹z
           Case vbKeyReturn
                                If PB_ENT.Visible Then                      'y“ü—Íz
                                '   Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                    Call PB_ENT_Click
                                '   Dummy.Enabled = False
                                End If
           Case Else:           Exit Function
    End Select
'
    Key_Acc = True
End Function
'+----------------------------------------------------------+
'+      ƒtƒH[ƒ€    s QueryUnloadƒCƒxƒ“ƒg t                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub
'+---------------------+
'+    I  —¹  ƒL [     +
'+---------------------+
Private Sub PB_END_Click()
    '+---------------------+
    '+   ÃŞ°ÀÍŞ°½Ø’f
    '+---------------------+
    Call CisDB.DBDISConnect
    '+---------------------+
    '+   Ì«°ÑµÌŞ¼Şª¸ÄÁ‹
    '+---------------------+
    Call CisFrm.UnLoadDisp
    End
End Sub
'+----------------------------+
'+    “ü  —Í  ƒL [(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call ReturnPress
    Dummy.Enabled = False
End Sub
'+----------------------------+
'+    æ  Á  ƒL [(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    Call CisFrm.MousePT(11)
    If ProcHB = "B1" Or _
       ProcHB = "B2" Then
        SyoriKB = ""
        Call SyoriName(SyoriKB)
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B3" Then
        Call CisVsGrid1.vsColor(Á‹ŒÂ•Ê)
        Select Case SyoriKB
            Case "ADD": ProcHB = "B1"
            Case "MNT": ProcHB = "B2"
        End Select
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        Call HeadBodyClear("B", "B3")
        If SyoriKB = "MNT" Then SendKeys "{TAB}"
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    V@‹K  ƒL [(F1)        +
'+----------------------------+
Private Sub PB_ADD_Click()
    If SyoriKB <> "ADD" Then
        ProcHB = "B1"
        Call DispChange(ProcHB)
        SyoriKB = "ADD"
        Call SyoriName(SyoriKB)
        Call HeadBodyClear("B", "B1")
        Call HeadBodyClear("B", "B3")
    End If
End Sub
'+----------------------------+
'+    C  ³  ƒL [(F2)        +
'+----------------------------+
Private Sub PB_MNT_Click()
    If SyoriKB <> "MNT" Then
        ProcHB = "B2"
        Call DispChange(ProcHB)
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
        SendKeys "{TAB}"
    End If
End Sub
'+----------------------------+
'+    í  œ  ƒL [(F3)        +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
        ProcHB = "B2"
        Call DispChange(ProcHB)
        SyoriKB = "DEL"
        Call SyoriName(SyoriKB)
        SendKeys "{TAB}"
    End If
End Sub
'+----------------------------+
'+    ˆó@ü  ƒL [(F8) +
'+----------------------------+
Private Sub PB_PRT_Click()
    
    cpYm = RTrim(iH1_SYm)
'ƒŠƒXƒgo—ÍæŠm”F
    With CisSimplePrint
         .PrintTitle = "x•¥ˆê——•\"
         .SaveKey = gRegKey
         .Preservation = "COK0010"
         If Not .CheckPrtCsv Then Exit Sub
    
         Select Case .PrintStatus
            Case 4: cpPrtKb = 0
            Case 5: cpPrtKb = 1
            Case 3: cpPrtKb = 2
            Case Else: Exit Sub
         End Select
         
    End With
'ˆóüˆ—
    Call COK0010LPrint
    Call CisSimplePrint.CheckFormUnload
End Sub
'+----------------------------+
'+    ŒŸ  õ  ƒL [(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    Select Case PB_Look.Tag
        Case "iB1_Torcd":  Call Look_Torcd
        Case "iB3_Paykb1": Call Look_Kubun(1)
        Case "iB3_Paykb2": Call Look_Kubun(2)
        Case "iB3_SYmd1":  iB3_SYmd(1).ShowCalender
        Case "iB3_SYmd2":  iB3_SYmd(2).ShowCalender
    End Select
End Sub
'+----------------------------+
'+       Œ_–ñæŒŸõ
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 1
    RV_TorKb1 = 1
'
    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing
'
    If RV_Rtn Then
        iB1_Torcd = RTrim(RV_TorcdK)
        Call iB1_Torcd_LostFocus
    End If
End Sub
'+----------------------------+
'+       –¼ÌŒŸõ
'+----------------------------+
Private Sub Look_Kubun(Index As Integer)
    RV_Left = 0
    RV_Top = 0
    RVI_Kubun = "”ƒŠ|x•¥‹æ•ª"
'
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing
'
    If RV_Rtn Then
        iB3_Paykb(Index) = RV_Code
        B3lb_Paynm(Index) = RV_Kubun
    End If
End Sub
'+-----------------------------+
'+      iB1_Torcd
'+-----------------------------+
Private Sub iB1_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B1lb_Tornm = ""
    If RTrim$(iB1_Torcd) <> "" Then
        If TorNmGet(iB1_Torcd, 5) Then
            B1lb_Tornm = TRM.—ªÌ
        End If
    End If
End Sub
'+-----------------------------+
'+      iB3_Paykb
'+-----------------------------+
Private Sub iB3_Paykb_GotFocus(Index As Integer)
    PB_Look.Tag = ActiveControl.Name & Format(Index, "0")
    PB_Look.Visible = True
End Sub
Private Sub iB3_Paykb_LostFocus(Index As Integer)
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
'
    B3lb_Paynm(Index) = ""
    If RTrim$(iB3_Paykb(Index)) <> "" Then
        If MeisyoGet("”ƒŠ|x•¥‹æ•ª", iB3_Paykb(Index)) Then
            B3lb_Paynm(Index) = SYM_Meisyo
        '   >> x•¥“ú‚ª–¢“ü—Í‚Í‰Šú¾¯Ä
            If Trim$(iB3_SYmd(Index)) = "" Then
                Select Case SYM.”š2
                    Case 0:    wAddSu = SYM.”š1 + 1: wDay = 1
                    Case Else: wAddSu = SYM.”š1:     wDay = SYM.”š2
                End Select
                wYmd = iH1_SYm & Format(wDay, "00")
                If IsDate(Format(wYmd, "@@@@/@@/@@")) Then
                    With CisFun
                        .Date_RsFormat = fYMD
                        .Date_Add wYmd, M, wAddSu
                        wYmd = Mid(.Date_Result, 1, 8)
                        If SYM.”š2 = 0 Then
                            .Date_Add wYmd, D, -1
                            wYmd = Mid(.Date_Result, 1, 8)
                        End If
                        iB3_SYmd(Index) = wYmd
                    End With
                End If
            End If
        End If
    End If
End Sub
'+-----------------------------+
'+      iB3_Ritu
'+-----------------------------+
Private Sub iB3_Ritu_LostFocus(Index As Integer)
    If CisFun.Val2(iB3_Ritu(Index)) <> 0 Then
    '   >> x•¥‹àŠz‚ÌŒvZ
        wKin = CisFun.Val2(B3lb_Sekyu) * (CisFun.Val2(iB3_Ritu(Index)) / 100)
        iB3_SKin(Index) = Format(wKin, "#,###")
    '
        If CisFun.Val2(iB3_Ritu(1)) + CisFun.Val2(iB3_Ritu(2)) = 100 Then
            If CisFun.Val2(iB3_SKin(1)) + CisFun.Val2(iB3_SKin(2)) <> CisFun.Val2(B3lb_Sekyu) Then
                If CisFun.Val2(iB3_Ritu(1)) < CisFun.Val2(iB3_Ritu(2)) Then
                    wKin = CisFun.Val2(B3lb_Sekyu) - CisFun.Val2(iB3_SKin(2))
                    iB3_SKin(1) = Format(wKin, "#,###")
                Else
                    wKin = CisFun.Val2(B3lb_Sekyu) - CisFun.Val2(iB3_SKin(1))
                    iB3_SKin(2) = Format(wKin, "#,###")
                End If
            End If
        End If
    End If
End Sub
'+-----------------------------+
'+      iB3_SYmd
'+-----------------------------+
Private Sub iB3_SYmd_GotFocus(Index As Integer)
    PB_Look.Tag = ActiveControl.Name & Format(Index, "0")
    PB_Look.Visible = True
End Sub
Private Sub iB3_SYmd_LostFocus(Index As Integer)
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+      iB3_SRKin
'+-----------------------------+
Private Sub iB3_SRKin_LostFocus()
'   >> d“üÁ”ïÅ‚ÌŒvZ
    wKin = CisFun.Val2(iB3_SRKin) * (wZRitu / 100)
    Select Case wZHasu
        Case 0: wKin = CisFun.HasuProc(wKin, ØÌ‚Ä, 0)
        Case 1: wKin = CisFun.HasuProc(wKin, Øã‚°, 0)
        Case 2: wKin = CisFun.HasuProc(wKin, lÌŒÜ“ü, 0)
    End Select
    B3lb_SRZei = Format(wKin, "#,###")
'   >> ‹àŠzŒvZ
    Call CompKin
End Sub
'+-----------------------------+
'+      iB3_SSKin
'+-----------------------------+
Private Sub iB3_SSKin_LostFocus()
'   >> ‘ŠEÁ”ïÅ‚ÌŒvZ
    wKin = CisFun.Val2(iB3_SSKin) * (wZRitu / 100)
    Select Case wZHasu
        Case 0: wKin = CisFun.HasuProc(wKin, ØÌ‚Ä, 0)
        Case 1: wKin = CisFun.HasuProc(wKin, Øã‚°, 0)
        Case 2: wKin = CisFun.HasuProc(wKin, lÌŒÜ“ü, 0)
    End Select
    B3lb_SSZei = Format(wKin, "#,###")
'   >> ‹àŠzŒvZ
    Call CompKin
End Sub
'+-----------------------------+
'+      iB3_SHKin
'+-----------------------------+
Private Sub iB3_SHKin_LostFocus()
'   >> ‹àŠzŒvZ
    Call CompKin
End Sub
'+-----------------------------+
'+      ‹àŠzŒvZF¿‹•x•¥
'+-----------------------------+
Private Sub CompKin()
    If wKurikosi = 0 Then
        wKin = CisFun.Val2(iB3_SRKin) + CisFun.Val2(B3lb_SRZei) _
             - CisFun.Val2(iB3_SSKin) - CisFun.Val2(B3lb_SSZei)
        B3lb_Sekyu = Format(wKin, "#,###")
        B3lb_TKuri = ""
    Else
        wKin = CisFun.Val2(B3lb_ZKuri) _
             + CisFun.Val2(iB3_SRKin) + CisFun.Val2(B3lb_SRZei) _
             - CisFun.Val2(iB3_SSKin) - CisFun.Val2(B3lb_SSZei)
        B3lb_Sekyu = Format(wKin, "#,###")
        wKin = wKin - CisFun.Val2(iB3_SHKin)
        B3lb_TKuri = Format(wKin, "#,###")
    End If
'   >> x•¥ŒvZ
    Call iB3_Ritu_LostFocus(1)
    Call iB3_Ritu_LostFocus(2)
End Sub
'**************************************
'*                                    *
'*      ˆ—‹æ•ª‚É‚æ‚è‹æ•ª–¼•\¦        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
    Dim w_Color     As Long
    w_Color = &H80000008 '‰ŠúF¾¯Ä(•)
'
    Select Case SYKB
        Case "ADD"
            SyoriNM = "’Ç‰Á"
            K_Sykbnm.ForeColor = &HFFFF&   '(‰©F)
        Case "MNT"
            SyoriNM = "C³"
            K_Sykbnm.ForeColor = &HFF0000  '(ÂF)
        Case "DEL"
            SyoriNM = "íœ"
            K_Sykbnm.ForeColor = &HFF&     '(ÔF)
            w_Color = &HFF&
        Case "LOK"
            SyoriNM = "Æ‰ï"
            K_Sykbnm.ForeColor = &HFF00FF   '(ƒsƒ“ƒN)
            PB_ADD.Visible = False
            PB_MNT.Visible = False
            PB_DEL.Visible = False
        Case Else
            SyoriNM = ""
    End Select
    For Each gObj In Me.Controls
        If gObj.Name Like "i*" Then gObj.ForeColor = w_Color
        If gObj.Name Like "*Comb*" Then gObj.ForeColor = w_Color
    '>> * PF-Key •\¦İ’è *
        If SYKB <> "LOK" Then
            If gObj.Name = "PB_ADD" Or _
               gObj.Name = "PB_MNT" Or _
               gObj.Name = "PB_DEL" Then
                gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SYKB)
            End If
        End If
    Next gObj
'
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
'*      ƒ^ƒCƒgƒ‹ƒo[•ÒW        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*******************************
'*      ˆ—‹æ•ª–¼“_–Å§Œä      *
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
'*      HEAD/BODY ˆ—      *
'*                          *
'****************************
'****************************
'*      “ü—ÍØ‘Ö ˆ—        *
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
'
    H1Mode = False: H1Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
    B2Mode = False: B2Color = gPLostSel
    B3Mode = False: B3Color = gPLostSel
'
' ÀŞĞ°€–Ú(‰æ–Ê‚Ìƒ`ƒ‰ƒcƒL–h~)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1": H1Mode = True: H1Color = gPGotSel
       Case "B1": B1Mode = True: B1Color = gPGotSel
       Case "B2": B2Mode = True: B2Color = gPGotSel
       Case "B3": B3Mode = True: B3Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode: H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode: B1_Area1.BackColor = B1Color
    B2_Area1.Enabled = B2Mode: B2_Area1.BackColor = B2Color
    B3_Area1.Enabled = B3Mode: B3_Area1.BackColor = B3Color
    B3_Area2.Enabled = B3Mode: B3_Area2.BackColor = B3Color
'
    PB_CAN.Visible = Not H1Mode
    PB_ADD.Visible = B1Mode Or B2Mode
    PB_MNT.Visible = B1Mode Or B2Mode
    PB_DEL.Visible = B1Mode Or B2Mode
    PB_END.Visible = H1Mode
    PB_PRT.Visible = B2Mode
'
' ÀŞĞ°€–Ú(“ü—Í‹Ö~)
    Dummy.Enabled = False
'
End Function
'**************************************
'*      Head/Body ‰Šú‰»ˆ—           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String
'
    If ClrType = "ALL" Then wbc_Enm = HeadBody & "*" Else wbc_Enm = Trim(ClrType)
'
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then .Text = ""
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
'
    If HeadBody = "B" And ClrType = "ALL" Then
        CisVsGrid1.Clear
        Call CisVsGrid1.Clear
        Dsp_Kensu = ""
    End If
End Sub
'****************************
'*      “ü—Íˆ—(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        If SyoriKB = "MNT" Then SendKeys "{TAB}"
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        ProcHB = "B3"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B2" Then
        If Not Body2Chk Then GoTo ReturnPress_Ed
    ' << íœˆ— >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then
                RowColFlg = True    '### vsGrid1_RowColChange : –³Œø‚É‚·‚é ###
                Call CisVsGrid1.vsColor(Á‹ŒÂ•Ê)
                RowColFlg = False   '### vsGrid1_RowColChange : —LŒø‚É–ß‚· ###
            Else
                If CisFun.Val2(Dsp_Kensu) <> 0 Then
                    ProcHB = "B2"
                Else
                    ProcHB = "B1"
                    SyoriKB = "ADD"
                End If
                Call DispChange(ProcHB)
                Call SyoriName(SyoriKB)
                If SyoriKB = "DEL" Then SendKeys "{TAB}"
            End If
        Else
            ProcHB = "B3"
            Call DispChange(ProcHB)
        End If
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B3" Then
        If Not Body3Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Select Case SyoriKB
            Case "ADD": ProcHB = "B1"
            Case "MNT": ProcHB = "B2"
        End Select
        Call DispChange(ProcHB)
        If SyoriKB = "MNT" Then SendKeys "{TAB}"
        GoTo ReturnPress_Ed
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1Áª¯¸ˆ—
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
'
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
'   >> x•¥”NŒ
    If Trim$(iH1_SYm) = "" Then
        CisFun.MB_MSG(2) = "    x•¥”NŒ‚ğ“ü—Í‚µ‚Ä‰º‚³‚¢            "
        CisFun.MBOX
        iH1_SYm.SetFocus
        Exit Function
    End If
    If Not iH1_SYm.cChkResult Then
        CisFun.MB_MSG(2) = "    ³‚µ‚¢x•¥”NŒ‚ğ“ü—Í‚µ‚Ä‰º‚³‚¢            "
        CisFun.MBOX
        iH1_SYm.SetFocus
        Exit Function
    End If
'
    If DBInput Then
        SyoriKB = "MNT"
        ProcHB = "B2"
    Else
        SyoriKB = "ADD"
        ProcHB = "B1"
    End If
'
    Head1Chk = True
End Function
'****************************
'*      BODY1Áª¯¸ˆ—
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
'
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
'   >> Œ_–ñæ
    If RTrim$(iB1_Torcd) = "" Then
        CisFun.MB_MSG(2) = "    Œ_–ñæ‚ğ“ü—Í‚µ‚Ä‰º‚³‚¢            "
        CisFun.MBOX
        iB1_Torcd.SetFocus
        Exit Function
    End If
    gSL_Select = ""
    gSL_Select = gSL_Select & "select ’÷“ú"
    gSL_Select = gSL_Select & "  from æˆøæƒ}ƒXƒ^"
    gSL_Select = gSL_Select & " where æˆøæCD = '" & RTrim$(iB1_Torcd) & "'"
    gSL_Select = gSL_Select & "   and æˆøæ‹æ•ª = 1"
    If Not TRMRead(gSL_Select, 1) Then
        CisFun.MB_MSG(2) = "    Œ_–ñæ@–¢“o˜^Iiæˆøæƒ}ƒXƒ^j            "
        CisFun.MBOX
        iB1_Torcd.SetFocus
        Exit Function
    End If
    wShime = TRM.’÷“ú
'
    If DBInput(iB1_Torcd) Then
        CisFun.MB_MSG(2) = "    ŠY“–ƒf[ƒ^@“o˜^ÏIi”ƒŠ|x•¥ƒe[ƒuƒ‹j            "
        CisFun.MBOX
        Exit Function
    End If
'   >> ŒJ‰z”»’è
    If wKurikosi = 1 Then
    '   >> ”NŒ‚Ì¾¯Ä•‘OŒŒJ‰z‚Ìõˆø
        With CisFun
            .Date_RsFormat = fYM
            .Date_Add iH1_SYm, M, -1
            wZenYm = Mid(.Date_Result, 1, 6)
        End With
        gSL_Select = ""
        gSL_Select = gSL_Select & "select ŒJ‰z‹àŠz"
        gSL_Select = gSL_Select & "  from ”ƒŠ|x•¥ƒe[ƒuƒ‹"
        gSL_Select = gSL_Select & " where ÀÑ”NŒ = '" & Trim$(wZenYm) & "'"
        gSL_Select = gSL_Select & "   and Œ_–ñæ = '" & RTrim$(iB1_Torcd) & "'"
        If KSHRead(gSL_Select, 1) Then
            B3lb_ZKuri = Format(KSH.ŒJ‰z‹àŠz, "#,###")
        End If
    End If
'
    Body1Chk = True
End Function
'****************************
'*      BODY2Áª¯¸ˆ—
'****************************
Private Function Body2Chk() As Boolean
    Body2Chk = False
'
    With VsGrid1
        If Trim(.TextMatrix(.Row, CisVsGrid1.FixedGet(1))) = "" Then Exit Function
'
        RowColFlg = True    '### vsGrid1_RowColChange : –³Œø‚É‚·‚é ###
        Select Case SyoriKB
            Case "MNT": Call CisVsGrid1.vsColor(ŒÂ•Ê, &HC0FFC0)
            Case "DEL": Call CisVsGrid1.vsColor(ŒÂ•Ê, &HFFC0FF)
        End Select
        RowColFlg = False   '### vsGrid1_RowColChange : —LŒø‚É–ß‚· ###
    End With
'
    Body2Chk = True
End Function
'****************************
'*      Body3Áª¯¸ˆ—
'****************************
Private Function Body3Chk() As Boolean
    Body3Chk = False
'
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
'
    For wIndex = 1 To 2
    '   >> “ü—Í”»’è
        If RTrim$(iB3_Paykb(wIndex)) <> "" Or _
           CisFun.Val2(iB3_Ritu(wIndex)) <> 0 Or _
           CisFun.Val2(iB3_SKin(wIndex)) <> 0 Or _
           CisFun.Val2(iB3_Tesu(wIndex)) <> 0 Or _
           Trim$(iB3_SYmd(wIndex)) <> "" Then
'
        '   >> x•¥‹æ•ª
            If RTrim$(iB3_Paykb(wIndex)) = "" Then
                CisFun.MB_MSG(2) = "    x•¥‹æ•ª‚ğ“ü—Í‚µ‚Ä‰º‚³‚¢            "
                CisFun.MBOX
                iB3_Paykb(wIndex).SetFocus
                Exit Function
            End If
            If RTrim$(iB3_Paykb(wIndex)) <> "" Then
                If Not MeisyoGet("”ƒŠ|x•¥‹æ•ª", iB3_Paykb(wIndex)) Then
                    CisFun.MB_MSG(2) = "    x•¥‹æ•ª@–¢“o˜^Ii–¼Ìƒ}ƒXƒ^j            "
                    CisFun.MBOX
                    iB3_Paykb(wIndex).SetFocus
                    Exit Function
                End If
            End If
        '   >> x•¥‹àŠz
            If CisFun.Val2(iB3_SKin(wIndex)) = 0 Then
                CisFun.MB_MSG(2) = "    x•¥‹àŠz‚ğ“ü—Í‚µ‚Ä‰º‚³‚¢            "
                CisFun.MBOX
                iB3_SKin(wIndex).SetFocus
                Exit Function
            End If
        '   >> x•¥“ú
            If Trim$(iB3_SYmd(wIndex)) = "" Then
                CisFun.MB_MSG(2) = "    x•¥“ú‚ğ“ü—Í‚µ‚Ä‰º‚³‚¢            "
                CisFun.MBOX
                iB3_SYmd(wIndex).SetFocus
                Exit Function
            End If
            If Not iB3_SYmd(wIndex).cChkResult Then
                CisFun.MB_MSG(2) = "    ³‚µ‚¢x•¥“ú‚ğ“ü—Í‚µ‚Ä‰º‚³‚¢            "
                CisFun.MBOX
                iB3_SYmd(wIndex).SetFocus
                Exit Function
            End If
'
        End If
    Next wIndex
'   >> —¦
    If CisFun.Val2(iB3_Ritu(1)) <> 0 And _
       CisFun.Val2(iB3_Ritu(2)) <> 0 Then
        If CisFun.Val2(iB3_Ritu(1)) + CisFun.Val2(iB3_Ritu(2)) <> 100 Then
            CisFun.MB_MSG(2) = "    —¦‚Ì‡Œv‚‚P‚O‚O“            "
            CisFun.MBOX
            iB3_Ritu(1).SetFocus
            Exit Function
        End If
    End If
'   >> x•¥‹àŠz
    If CisFun.Val2(iB3_SKin(1)) + CisFun.Val2(iB3_SKin(2)) <> CisFun.Val2(B3lb_Sekyu) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    x•¥‹àŠz‚Ì‡Œv‚“–Œ¿‹‹àŠz            "
            .MB_MSG(3) = "          ‚æ‚ë‚µ‚¢‚Å‚·‚©H      "
            .MB_Title = "Šm”F‰æ–Ê"
            .MB_Button = OK_CAN
            If Not .MBOX Then Exit Function
        End With
    End If
'
    Body3Chk = True
End Function
'***********************************************************************************
'*
'*       ƒf[ƒ^ƒx[ƒX“ü—Í • •\¦
'*
'***********************************************************************************
Private Function DBInput(Optional Torcd As String = "") As Boolean
    DBInput = False
'
    On Error GoTo DBInput_Err:
'
    Call SqlSelectSet(Torcd)
    If KSHRead(gSL_Select, 0) Then
        If Torcd = "" Then
            Call Grid1Set
        End If
        DBInput = True
    End If
    Call KSHClose
'
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'****************************
'*      SQL•¶¾¯Äˆ—
'****************************
Private Sub SqlSelectSet(Optional Torcd As String = "")
'
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  ksh.*"
    gSL_Select = gSL_Select & ", isnull(trm.—ªÌ, '') Œ_–ñæ–¼"
    gSL_Select = gSL_Select & ", isnull(sy1.’l–¼Ì, '') x•¥‹æ•ª–¼1"
    gSL_Select = gSL_Select & ", isnull(sy2.’l–¼Ì, '') x•¥‹æ•ª–¼2"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  ”ƒŠ|x•¥ƒe[ƒuƒ‹ ksh"
    gSL_Select = gSL_Select & "  left outer join æˆøæƒ}ƒXƒ^ trm"
    gSL_Select = gSL_Select & "    on trm.æˆøæCD = ksh.Œ_–ñæ"
    gSL_Select = gSL_Select & "   and trm.æˆøæ‹æ•ª = 1"
    gSL_Select = gSL_Select & "  left outer join –¼Ìƒ}ƒXƒ^ sy1"
    gSL_Select = gSL_Select & "    on sy1.‹æ•ª–¼Ì = '”ƒŠ|x•¥‹æ•ª'"
    gSL_Select = gSL_Select & "   and sy1.‹æ•ªƒ^ƒCƒv = ''"
    gSL_Select = gSL_Select & "   and sy1.’l = ksh.x•¥‹æ•ª1"
    gSL_Select = gSL_Select & "  left outer join –¼Ìƒ}ƒXƒ^ sy2"
    gSL_Select = gSL_Select & "    on sy2.‹æ•ª–¼Ì = '”ƒŠ|x•¥‹æ•ª'"
    gSL_Select = gSL_Select & "   and sy2.‹æ•ªƒ^ƒCƒv = ''"
    gSL_Select = gSL_Select & "   and sy2.’l = ksh.x•¥‹æ•ª2"
    gSL_Select = gSL_Select & " where ksh.ÀÑ”NŒ = '" & Trim$(iH1_SYm) & "'"
    If RTrim$(Torcd) <> "" Then
        gSL_Select = gSL_Select & "   and ksh.Œ_–ñæ = '" & RTrim$(Torcd) & "'"
    End If
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  ksh.Œ_–ñæ"
'
End Sub
'**************************************
'*                                    *
'*          ƒf[ƒ^ƒx[ƒXXV           *
'*                                    *
'**************************************
'+----------------------------+
'+      í œ ˆ —            +
'+----------------------------+
Private Function DBDelete()
    DBDelete = False
'
'  ( Šm”FƒƒbƒZ[ƒW )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = "     íœˆ—‚ğs‚¢‚Ü‚·            "
        .MB_MSG(3) = "     ‚æ‚ë‚µ‚¢‚Å‚·‚©H      "
        .MB_Title = "Šm”F‰æ–Ê"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
'
    On Error GoTo DBDelete_Err
    Call CisDB.DBTran(TransBegin)   ' Ä×İ»Ş¸¼®İŠJn
'
'   >> ”ƒŠ|x•¥ƒe[ƒuƒ‹íœ
    If Not DBDelete_KSH Then GoTo DBDelete_Err
'   >> Œ”ƒJƒEƒ“ƒg
    iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")
'
    Call CisDB.DBTran(TransCommit)  ' Ä×İ»Ş¸¼®İ I—¹
    On Error GoTo 0
'
'   >> GridÄ•\¦
    Call vsGrid1_ReSetDisp
'
    DBDelete = True
    Exit Function
DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    ’Ç‰Á¥XV ˆ—     +
'+----------------------+
Private Function DBPut()
    DBPut = False
'
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    With CisFun
        .MB_Lines = 4
        Select Case SyoriKB
            Case "ADD": .MB_MSG(1) = "     ’Ç‰Áˆ—‚ğs‚¢‚Ü‚·            "
            Case "MNT": .MB_MSG(2) = "     C³ˆ—‚ğs‚¢‚Ü‚·            "
        End Select
        .MB_MSG(3) = "     ‚æ‚ë‚µ‚¢‚Å‚·‚©H      "
        .MB_Title = "Šm”F‰æ–Ê"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
'
    On Error GoTo DBPut_Err
    Call CisDB.DBTran(TransBegin)   ' Ä×İ»Ş¸¼®İŠJn
'
    Select Case SyoriKB
        Case "ADD"
    '   ( ’Ç‰Áˆ— )
            Call ItemsClearKSH
            Call DBEdit_KSH(1)
            Call KSHInsert
            iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
        Case "MNT"
    '   ( XVˆ— )
            Call DBEdit_KSH(2)
            Call DBUpdate_KSH
            iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
    End Select
'
    Call CisDB.DBTran(TransCommit)  ' Ä×İ»Ş¸¼®İ I—¹
    On Error GoTo 0
'
'   >> GridÄ•\¦
    Call vsGrid1_ReSetDisp
'
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'==========================================================='
'       íœˆ— - ”ƒŠ|x•¥ƒe[ƒuƒ‹
'==========================================================='
Private Function DBDelete_KSH(Optional DBNo As Integer = 0) As Boolean
    DBDelete_KSH = False
'
    gSL_Select = ""
    gSL_Select = gSL_Select & "delete"
    gSL_Select = gSL_Select & "  from ”ƒŠ|x•¥ƒe[ƒuƒ‹"
    gSL_Select = gSL_Select & " where ÀÑ”NŒ = '" & Trim$(iH1_SYm) & "'"
    gSL_Select = gSL_Select & "   and Œ_–ñæ = '" & RTrim$(iB1_Torcd) & "'"
'
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
'
    DBDelete_KSH = True
End Function
'==========================================================='
'       XVˆ— - ”ƒŠ|x•¥ƒe[ƒuƒ‹
'==========================================================='
Private Function DBUpdate_KSH(Optional DBNo As Integer = 0) As Boolean
    DBUpdate_KSH = False
'
    With KSH
        gSL_Select = ""
        gSL_Select = gSL_Select & "update ”ƒŠ|x•¥ƒe[ƒuƒ‹"
        gSL_Select = gSL_Select & "   set d“ü‹àŠz = " & .d“ü‹àŠz & ""
        gSL_Select = gSL_Select & "     , d“üÁ”ïÅ = " & .d“üÁ”ïÅ & ""
        gSL_Select = gSL_Select & "     , ‘ŠE‹àŠz = " & .‘ŠE‹àŠz & ""
        gSL_Select = gSL_Select & "     , ‘ŠEÁ”ïÅ = " & .‘ŠEÁ”ïÅ & ""
        gSL_Select = gSL_Select & "     , ¿‹‹àŠz = " & .¿‹‹àŠz & ""
        gSL_Select = gSL_Select & "     , x•¥‹àŠz = " & .x•¥‹àŠz & ""
        gSL_Select = gSL_Select & "     , ŒJ‰z‹àŠz = " & .ŒJ‰z‹àŠz & ""
        gSL_Select = gSL_Select & "     , x•¥‹æ•ª1 = '" & .x•¥‹æ•ª1 & "'"
        gSL_Select = gSL_Select & "     , x•¥—¦1 = " & .x•¥—¦1 & ""
        gSL_Select = gSL_Select & "     , x•¥“ú1 = '" & .x•¥“ú1 & "'"
        gSL_Select = gSL_Select & "     , x•¥è”—¿1 = " & .x•¥è”—¿1 & ""
        gSL_Select = gSL_Select & "     , x•¥‹àŠz1 = " & .x•¥‹àŠz1 & ""
        gSL_Select = gSL_Select & "     , x•¥‹æ•ª2 = '" & .x•¥‹æ•ª2 & "'"
        gSL_Select = gSL_Select & "     , x•¥—¦2 = " & .x•¥—¦2 & ""
        gSL_Select = gSL_Select & "     , x•¥“ú2 = '" & .x•¥“ú2 & "'"
        gSL_Select = gSL_Select & "     , x•¥è”—¿2 = " & .x•¥è”—¿2 & ""
        gSL_Select = gSL_Select & "     , x•¥‹àŠz2 = " & .x•¥‹àŠz2 & ""
        gSL_Select = gSL_Select & "     , XV“ú = getdate()"
        gSL_Select = gSL_Select & "     , XV’[–– = substring(host_name(), 1, 20)"
        gSL_Select = gSL_Select & "     , XVÒ = '" & RTrim$(.XVÒ) & "'"
        gSL_Select = gSL_Select & " where ÀÑ”NŒ = '" & .ÀÑ”NŒ & "'"
        gSL_Select = gSL_Select & "   and Œ_–ñæ = '" & .Œ_–ñæ & "'"
    End With
'
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
'
    DBUpdate_KSH = True
End Function
'==========================================================='
'       ”ƒŠ|x•¥ƒe[ƒuƒ‹•ÒW
'==========================================================='
Private Function DBEdit_KSH(KB As Byte) As Boolean
    DBEdit_KSH = False
'
    With KSH
        .ÀÑ”NŒ = RTrim$(iH1_SYm)
        .Œ_–ñæ = RTrim$(iB1_Torcd)
        .‘OŒŒJ‰z = CisFun.Val2(B3lb_ZKuri)
        .d“ü‹àŠz = CisFun.Val2(iB3_SRKin)
        .d“üÁ”ïÅ = CisFun.Val2(B3lb_SRZei)
        .‘ŠE‹àŠz = CisFun.Val2(iB3_SSKin)
        .‘ŠEÁ”ïÅ = CisFun.Val2(B3lb_SSZei)
        .¿‹‹àŠz = CisFun.Val2(B3lb_Sekyu)
        .x•¥‹àŠz = CisFun.Val2(iB3_SHKin)
        .ŒJ‰z‹àŠz = CisFun.Val2(B3lb_TKuri)
        .x•¥‹æ•ª1 = RTrim$(iB3_Paykb(1))
        .x•¥—¦1 = CisFun.Val2(iB3_Ritu(1))
        .x•¥“ú1 = Trim$(iB3_SYmd(1))
        .x•¥è”—¿1 = CisFun.Val2(iB3_Tesu(1))
        .x•¥‹àŠz1 = CisFun.Val2(iB3_SKin(1))
        .x•¥‹æ•ª2 = RTrim$(iB3_Paykb(2))
        .x•¥—¦2 = CisFun.Val2(iB3_Ritu(2))
        .x•¥“ú2 = Trim$(iB3_SYmd(2))
        .x•¥è”—¿2 = CisFun.Val2(iB3_Tesu(2))
        .x•¥‹àŠz2 = CisFun.Val2(iB3_SKin(2))
        .’÷“ú = wShime
        Select Case KB
            Case 1: .ì¬Ò = gTanto
            Case 2: .XVÒ = gTanto
        End Select
    End With
'
    DBEdit_KSH = True
End Function
'
'+==========================================================+
'+      ƒOƒŠƒbƒhˆ—                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   Œ©o‚µ¾¯Ä & ‰Šúó‘Ô¾¯Ä
'+----------------------------------------------------------+
Private Sub GridInit1()
    With CisVsGrid1
'@¸ŞØ¯ÄŞ‰Šú‰»ˆ—
'       .Clear
'       Call .FixedSet(Œ©o‚µˆÊ’uİ’è,—ñˆÊ’uİ’è,Ï°¼Şˆ—,
'                      ¾ÙŒ…”,—ñ”Ô†,Œ©o‚µ“à—e,•ª—ŞÀ²Ìß)
       'ReDim SColZ(21)
        Call .FixedSet(’†’†, ‰E’†, ‚ ‚è, 3, 0, "No")
        Call .FixedSet(’†’†, ¶’†, ‚ ‚è, 28, 1, "Œ_–ñæ")
        Call .FixedSet(’†’†, ’†’†, ‚ ‚è, 10, 2, "¿‹‘‡‚")
        Call .FixedSet(’†’†, ‰E’†, ‚ ‚è, 11, 3, "‘OŒŒJ‰z")
        Call .FixedSet(’†’†, ‰E’†, ‚ ‚è, 11, 4, "d“ü‹àŠz")
        Call .FixedSet(’†’†, ‰E’†, ‚ ‚è, 10, 5, "d“üÁ”ïÅ")
        Call .FixedSet(’†’†, ‰E’†, ‚ ‚è, 11, 6, "‘ŠE‹àŠz")
        Call .FixedSet(’†’†, ‰E’†, ‚ ‚è, 10, 7, "‘ŠEÁ”ïÅ")
        Call .FixedSet(’†’†, ‰E’†, ‚ ‚è, 11, 8, "¿‹‹àŠz")
        Call .FixedSet(’†’†, ‰E’†, ‚ ‚è, 11, 9, "x•¥‹àŠz")
        Call .FixedSet(’†’†, ‰E’†, ‚ ‚è, 11, 10, "ŒJ‰z‹àŠz")
        Call .FixedSet(’†’†, ¶’†, ‚ ‚è, 10, 11, "x•¥•û–@1")
        Call .FixedSet(’†’†, ‰E’†, ‚ ‚è, 3, 12, "—¦1")
        Call .FixedSet(’†’†, ‰E’†, ‚ ‚è, 11, 13, "x•¥‹àŠz1")
        Call .FixedSet(’†’†, ‰E’†, ‚ ‚è, 7, 14, "è”—¿1")
        Call .FixedSet(’†’†, ’†’†, ‚ ‚è, 10, 15, "x•¥“ú1")
        Call .FixedSet(’†’†, ¶’†, ‚ ‚è, 10, 16, "x•¥•û–@2")
        Call .FixedSet(’†’†, ‰E’†, ‚ ‚è, 3, 17, "—¦2")
        Call .FixedSet(’†’†, ‰E’†, ‚ ‚è, 11, 18, "x•¥‹àŠz2")
        Call .FixedSet(’†’†, ‰E’†, ‚ ‚è, 7, 19, "è”—¿2")
        Call .FixedSet(’†’†, ’†’†, ‚ ‚è, 10, 20, "x•¥“ú2")
        Call .FixedSet(’†’†, ‰E’†, ‚ ‚è, 4, 21, "’÷“ú")
'@¸ŞØ¯ÄŞ‘I‘ğˆ——p@‘Ş”ğ´Ø±
        Call .FixedSet(’†’†, ¶’†, ‚ ‚è, 0, 22, "XŒ_–ñæ")
        Call .FixedSet(’†’†, ¶’†, ‚ ‚è, 0, 23, "XŒ_–ñæ–¼")
        Call .FixedSet(’†’†, ¶’†, ‚ ‚è, 0, 24, "Xx•¥‹æ•ª1")
        Call .FixedSet(’†’†, ¶’†, ‚ ‚è, 0, 25, "Xx•¥‹æ•ª–¼1")
        Call .FixedSet(’†’†, ¶’†, ‚ ‚è, 0, 26, "Xx•¥“ú1")
        Call .FixedSet(’†’†, ¶’†, ‚ ‚è, 0, 27, "Xx•¥‹æ•ª2")
        Call .FixedSet(’†’†, ¶’†, ‚ ‚è, 0, 28, "Xx•¥‹æ•ª–¼2")
        Call .FixedSet(’†’†, ¶’†, ‚ ‚è, 0, 29, "Xx•¥“ú2")
        Call .FixedSet(’†’†, ¶’†, ‚ ‚è, 0, 30, "Select")
'  •\¦ÊßÀ°İ¾¯Ä
'       Call .PatanSet(ÊßÀ°İNO,•\¦ƒpƒ^[ƒ“)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21")
'  •ÒW“à—eƒZƒbƒg
'       Call .EditSet(Ã·½ÄµÌŞ¼Şª¸Ä,¶×ÑNO)
'       Call .EditSet(iH_Text1, 1)
        .InitDisp
    End With
End Sub
'****************************************************************************************'
'*      ƒOƒŠƒbƒh“à—eƒZƒbƒg
'****************************************************************************************'
Private Sub Grid1Set()
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
'
'   >> ’Šoˆ— ‰Šú¾¯Ä
    Ttl_Kensu = CisDB.RecordCount       ' ’ŠoŒ”(‡Œv)‘Ş”ğ
    Dsp_Kensu = Format(Ttl_Kensu, "#,###")
'   >> ¸ŞØ¯ÄŞsÄ¾¯Ä & —ñˆÊ’uæ“¾
    With VsGrid1
       .Redraw = False
        If Ttl_Kensu + 1 < CisVsGrid1.DispMax Then
            .Rows = CisVsGrid1.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
'
    ID = 0
    Do Until Not KSH_RDSTS
        ID = ID + 1
'
        Call DBInfoToGrid1(ID)
'
        Call KSHReadNext
    Loop
'
    With VsGrid1
        .Row = 1: .Col = 1: .ColSel = .Cols - 1
        .Redraw = True
    End With
End Sub
'+----------------------------------------------+
'+      ƒe[ƒuƒ‹‚æ‚èƒOƒŠƒbƒh‚É“]‘—
'+----------------------------------------------+
Private Sub DBInfoToGrid1(SetRow As Long)
    With VsGrid1
    '>> ‡‚
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(0)) = Format(SetRow, "###")
    '>> Œ_–ñæ–¼
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(1)) = KSH.Œ_–ñæ & ":" & RTrim$(KSH.Œ_–ñæ–¼)
    '>> ¿‹‘NO
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(2)) = RTrim$(KSH.¿‹‘NO)
    '>> ‘OŒŒJ‰z
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(3)) = Format(KSH.‘OŒŒJ‰z, "#,###")
    '>> d“ü‹àŠz
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(4)) = Format(KSH.d“ü‹àŠz, "#,###")
    '>> d“üÁ”ïÅ
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(5)) = Format(KSH.d“üÁ”ïÅ, "#,###")
    '>> ‘ŠE‹àŠz
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(6)) = Format(KSH.‘ŠE‹àŠz, "#,###")
    '>> ‘ŠEÁ”ïÅ
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(7)) = Format(KSH.‘ŠEÁ”ïÅ, "#,###")
    '>> ¿‹‹àŠz
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(8)) = Format(KSH.¿‹‹àŠz, "#,###")
    '>> x•¥‹àŠz
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(9)) = Format(KSH.x•¥‹àŠz, "#,###")
    '>> ŒJ‰z‹àŠz
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(10)) = Format(KSH.ŒJ‰z‹àŠz, "#,###")
    '>> x•¥-1
        If RTrim$(KSH.x•¥‹æ•ª1) <> "" Then
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(11)) = KSH.x•¥‹æ•ª1 & ":" & RTrim$(KSH.x•¥‹æ•ª–¼1)
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(12)) = Format(KSH.x•¥—¦1, "###")
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(13)) = Format(KSH.x•¥‹àŠz1, "#,###")
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(14)) = Format(KSH.x•¥è”—¿1, "###")
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(15)) = Format(KSH.x•¥“ú1, "@@@@/@@/@@")
        End If
    '>> x•¥-2
        If RTrim$(KSH.x•¥‹æ•ª2) <> "" Then
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(16)) = KSH.x•¥‹æ•ª2 & ":" & RTrim$(KSH.x•¥‹æ•ª–¼2)
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(17)) = Format(KSH.x•¥—¦2, "###")
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(18)) = Format(KSH.x•¥‹àŠz2, "#,###")
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(19)) = Format(KSH.x•¥è”—¿2, "###")
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(20)) = Format(KSH.x•¥“ú2, "@@@@/@@/@@")
        End If
    '>> ’÷“ú
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(21)) = Format(KSH.’÷“ú, "##")
    '>> ‘Ş”ğ
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(22)) = RTrim$(KSH.Œ_–ñæ)
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(23)) = RTrim$(KSH.Œ_–ñæ–¼)
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(24)) = RTrim$(KSH.x•¥‹æ•ª1)
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(25)) = RTrim$(KSH.x•¥‹æ•ª–¼1)
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(26)) = Trim$(KSH.x•¥“ú1)
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(27)) = RTrim$(KSH.x•¥‹æ•ª2)
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(28)) = RTrim$(KSH.x•¥‹æ•ª–¼2)
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(29)) = Trim$(KSH.x•¥“ú2)
    End With
End Sub
'+---------------------------------------------------------------------------------------+
'+  ƒOƒŠƒbƒh‚æ‚è“ü—Íˆæ‚ÖƒZƒbƒg
'+---------------------------------------------------------------------------------------+
Private Sub GridToText()
    With VsGrid1
        iB1_Torcd = .TextMatrix(.Row, CisVsGrid1.FixedGet(22))
        B1lb_Tornm = .TextMatrix(.Row, CisVsGrid1.FixedGet(23))
    '
        iB3_SRKin = CisFun.Val2(.TextMatrix(.Row, CisVsGrid1.FixedGet(4)))
        iB3_SSKin = CisFun.Val2(.TextMatrix(.Row, CisVsGrid1.FixedGet(6)))
        iB3_SHKin = CisFun.Val2(.TextMatrix(.Row, CisVsGrid1.FixedGet(9)))
        B3lb_ZKuri = .TextMatrix(.Row, CisVsGrid1.FixedGet(3))
        B3lb_SRZei = .TextMatrix(.Row, CisVsGrid1.FixedGet(5))
        B3lb_SSZei = .TextMatrix(.Row, CisVsGrid1.FixedGet(7))
        B3lb_Sekyu = .TextMatrix(.Row, CisVsGrid1.FixedGet(8))
        B3lb_TKuri = .TextMatrix(.Row, CisVsGrid1.FixedGet(10))
        B3lb_SeiNo = .TextMatrix(.Row, CisVsGrid1.FixedGet(2))
    '
        iB3_Paykb(1) = .TextMatrix(.Row, CisVsGrid1.FixedGet(24))
        iB3_Ritu(1) = CisFun.Val2(.TextMatrix(.Row, CisVsGrid1.FixedGet(12)))
        iB3_SKin(1) = CisFun.Val2(.TextMatrix(.Row, CisVsGrid1.FixedGet(13)))
        iB3_Tesu(1) = CisFun.Val2(.TextMatrix(.Row, CisVsGrid1.FixedGet(14)))
        iB3_SYmd(1) = .TextMatrix(.Row, CisVsGrid1.FixedGet(26))
        B3lb_Paynm(1) = .TextMatrix(.Row, CisVsGrid1.FixedGet(25))
        iB3_Paykb(2) = .TextMatrix(.Row, CisVsGrid1.FixedGet(27))
        iB3_Ritu(2) = CisFun.Val2(.TextMatrix(.Row, CisVsGrid1.FixedGet(17)))
        iB3_SKin(2) = CisFun.Val2(.TextMatrix(.Row, CisVsGrid1.FixedGet(18)))
        iB3_Tesu(2) = CisFun.Val2(.TextMatrix(.Row, CisVsGrid1.FixedGet(19)))
        iB3_SYmd(2) = .TextMatrix(.Row, CisVsGrid1.FixedGet(29))
        B3lb_Paynm(2) = .TextMatrix(.Row, CisVsGrid1.FixedGet(28))
    End With
End Sub
'+---------------------------------------+
'+      ƒOƒŠƒbƒh    (Click)
'+---------------------------------------+
Private Sub VSGrid1_Click()
    If ProcHB = "B2" Then
        Call GridToText
    End If
End Sub
'+---------------------------------------+
'+      ƒOƒŠƒbƒh    (DblClick)
'+---------------------------------------+
Private Sub vsGrid1_DblClick()
    If ProcHB = "B2" Then
        Call ReturnPress
    End If
End Sub
'+---------------------------------------+
'+      ƒOƒŠƒbƒh    (KeyPress)
'+---------------------------------------+
Private Sub vsGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
       Call vsGrid1_DblClick
    End If
End Sub
'+---------------------------------------+
'+      ƒOƒŠƒbƒh    (RowColChange)
'+---------------------------------------+
Private Sub vsGrid1_RowColChange()
    If FormAct Then
    '>> Às”»’è .. False:Às
        Select Case RowColFlg
            Case False: Call VSGrid1_Click
        End Select
    End If
End Sub
'+---------------------------------------+
'+      ƒOƒŠƒbƒh Ä•\¦
'+---------------------------------------+
Private Sub vsGrid1_ReSetDisp()
    SaveRow = VsGrid1.Row
'
    RowColFlg = True    '### vsGrid1_RowColChange : –³Œø‚É‚·‚é ###
    Call HeadBodyClear("B")
    Call DBInput
    RowColFlg = False   '### vsGrid1_RowColChange : —LŒø‚É–ß‚· ###
'
    If SaveRow > VsGrid1.Rows - 1 Then
        SaveRow = SaveRow - 1
    End If
    If SyoriKB <> "ADD" Then
        VsGrid1.Row = SaveRow
    End If
    If VsGrid1.RowIsVisible(SaveRow) = False Then
        VsGrid1.TopRow = SaveRow
    End If
End Sub




