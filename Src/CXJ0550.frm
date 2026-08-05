VERSION 5.00
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{BA5D7819-2B67-11D4-806A-00000EA57F9E}#3.0#0"; "CISYMD~1.OCX"
Object = "{1BA4D453-0150-11CE-89B0-0000C037528B}#3.2#0"; "PDQCom32.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXJ0550 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "受注チケット入力"
   ClientHeight    =   11235
   ClientLeft      =   1800
   ClientTop       =   1515
   ClientWidth     =   15360
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00004000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin VB.PictureBox Picture3 
      Height          =   3375
      Left            =   12000
      ScaleHeight     =   3315
      ScaleWidth      =   1440
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   6015
      Visible         =   0   'False
      Width           =   1500
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   52
         Left            =   0
         Top             =   450
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_KiriF 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   84
            Top             =   90
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   54
         Left            =   0
         Top             =   1020
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_KiriF 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   85
            Top             =   90
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   55
         Left            =   0
         Top             =   1590
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_KiriF 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   86
            Top             =   90
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   70
         Left            =   0
         Top             =   2160
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_KiriF 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   87
            Top             =   90
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   71
         Left            =   0
         Top             =   2730
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_KiriF 
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   88
            Top             =   90
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   435
         Index           =   72
         Left            =   0
         Top             =   0
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   767
         BackColor       =   14737632
         ForeColor       =   16711680
         Caption         =   "打切Ｆ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   435
         Index           =   73
         Left            =   750
         Top             =   0
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   767
         BackColor       =   14737632
         ForeColor       =   16711680
         Caption         =   "打切残"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   74
         Left            =   750
         Top             =   450
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_KiriZ 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   89
            Top             =   90
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   75
         Left            =   750
         Top             =   1020
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_KiriZ 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   90
            Top             =   90
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   76
         Left            =   750
         Top             =   1590
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_KiriZ 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   91
            Top             =   90
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   77
         Left            =   750
         Top             =   2160
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_KiriZ 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   92
            Top             =   90
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   78
         Left            =   750
         Top             =   2730
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_KiriZ 
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   93
            Top             =   90
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   5295
      Left            =   -1800
      ScaleHeight     =   5235
      ScaleWidth      =   1875
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   1935
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   915
         Index           =   1
         Left            =   1200
         Top             =   3360
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "行"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText CisText4 
            Height          =   375
            Left            =   90
            TabIndex        =   81
            Top             =   420
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
            cFbComma        =   0
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   915
         Index           =   51
         Left            =   720
         Top             =   3360
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "CD"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText CisText3 
            Height          =   375
            Left            =   120
            TabIndex        =   80
            Top             =   420
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   57
         Left            =   0
         Top             =   450
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_UkeNo 
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   67
            Top             =   90
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   58
         Left            =   0
         Top             =   1020
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_UkeNo 
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   68
            Top             =   90
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   60
         Left            =   0
         Top             =   1590
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_UkeNo 
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   69
            Top             =   90
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   61
         Left            =   0
         Top             =   2160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_UkeNo 
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   71
            Top             =   90
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   62
         Left            =   0
         Top             =   2730
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_UkeNo 
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   72
            Top             =   90
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   435
         Index           =   56
         Left            =   0
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   767
         BackColor       =   14737632
         ForeColor       =   16711680
         Caption         =   "受付No"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   435
         Index           =   63
         Left            =   960
         Top             =   0
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   767
         BackColor       =   14737632
         ForeColor       =   16711680
         Caption         =   "CD"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   64
         Left            =   975
         Top             =   450
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_CD 
            Height          =   375
            Index           =   5
            Left            =   165
            TabIndex        =   73
            Top             =   90
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   65
         Left            =   975
         Top             =   1020
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_CD 
            Height          =   375
            Index           =   6
            Left            =   165
            TabIndex        =   74
            Top             =   90
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   67
         Left            =   975
         Top             =   1590
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_CD 
            Height          =   375
            Index           =   7
            Left            =   165
            TabIndex        =   75
            Top             =   90
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   68
         Left            =   975
         Top             =   2160
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_CD 
            Height          =   375
            Index           =   8
            Left            =   165
            TabIndex        =   76
            Top             =   90
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   555
         Index           =   69
         Left            =   975
         Top             =   2730
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin CisText_V60.CisText iB12_CD 
            Height          =   375
            Index           =   9
            Left            =   165
            TabIndex        =   77
            Top             =   90
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   915
         Index           =   49
         Left            =   0
         Top             =   3360
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "区"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText CisText1 
            Height          =   375
            Left            =   120
            TabIndex        =   78
            Top             =   420
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
            cFbComma        =   0
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   915
         Index           =   50
         Left            =   480
         Top             =   3360
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "#"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText CisText2 
            Height          =   375
            Left            =   90
            TabIndex        =   79
            Top             =   420
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
            cFbComma        =   0
         End
      End
      Begin Cis3D_v60.CIS3D B1_C_Back 
         Height          =   915
         Left            =   -480
         Top             =   4320
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "年　月　日　便"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText iB1_MM_ 
            Height          =   375
            Left            =   630
            TabIndex        =   95
            Top             =   420
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            cFormat         =   "00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText iB1_DD_ 
            Height          =   375
            Left            =   1110
            TabIndex        =   96
            Top             =   420
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            cFormat         =   "00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText iB1_Bin_ 
            Height          =   375
            Left            =   1605
            TabIndex        =   97
            Top             =   420
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText iB1_YY_ 
            Height          =   375
            Left            =   180
            TabIndex        =   98
            Top             =   420
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            cFormat         =   "00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
   End
   Begin VB.PictureBox H1_Area2 
      BackColor       =   &H00C0C0C0&
      Height          =   1065
      Left            =   2250
      ScaleHeight     =   1005
      ScaleWidth      =   2610
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   4320
      Width           =   2670
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   870
         Index           =   0
         Left            =   60
         Top             =   60
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   1535
         ForeColor       =   16711680
         Caption         =   "－ 便  "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   3
         cPositionX      =   120
         cPositionY      =   75
         Begin CisYMD_With_Btn.CisYMDwB iH1_Ymd 
            Height          =   345
            Left            =   90
            TabIndex        =   8
            Top             =   420
            Width           =   1950
            _ExtentX        =   2990
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Object.Height          =   345
            Object.Width           =   1950
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Cis3D_v60.CIS3D Back_Ymd 
            Height          =   300
            Left            =   510
            Top             =   75
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "納 入 日"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText iH1_Bin 
            Height          =   360
            Left            =   2055
            TabIndex        =   9
            Top             =   420
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            cFormat         =   "00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
   End
   Begin PdqcommLib.PDQComm PDQComm1 
      Height          =   480
      Left            =   5025
      TabIndex        =   63
      Top             =   375
      Width           =   480
      _Version        =   196610
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.01
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FastScroll      =   0   'False
      Object.Height          =   0
      Object.Width           =   0
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   7185
      ScaleHeight     =   735
      ScaleWidth      =   5265
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   1155
      Width           =   5325
      Begin CisBtn_60.CisBtn PB_RED 
         Height          =   585
         Left            =   60
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnCaption      =   "読取"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   8388736
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
         PFBackColor     =   8388736
      End
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   585
         Left            =   930
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   7
         BtnCaption      =   "新規"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   2655
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   3
         BtnCaption      =   "取消"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         Left            =   3510
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   2
         BtnCaption      =   "入力"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         Left            =   4365
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   1
         BtnCaption      =   "終了"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         Left            =   1800
         Top             =   60
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   11
         BtnCaption      =   "検索"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   9330
      ScaleHeight     =   1005
      ScaleWidth      =   2460
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   4320
      Width           =   2520
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   915
         Left            =   3960
         Top             =   -60
         Visible         =   0   'False
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "便"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText iB1_Bin 
            Height          =   375
            Left            =   0
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   420
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   915
         Index           =   5
         Left            =   1710
         Top             =   60
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "機番"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText iB1_Kibn 
            Height          =   375
            Left            =   150
            TabIndex        =   13
            Top             =   420
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
            cFbComma        =   0
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   915
         Index           =   4
         Left            =   1080
         Top             =   60
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "受入"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText iB1_Uke 
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   420
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   915
         Index           =   3
         Left            =   60
         Top             =   60
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "発注者"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
         Begin CisText_V60.CisText iB1_Hacyu 
            Height          =   375
            Left            =   135
            TabIndex        =   11
            Top             =   420
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   915
         Index           =   46
         Left            =   25230
         Top             =   615
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "受入"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
      Begin CisYMD_With_Btn.CisYMDwB iB1_Ymd 
         Height          =   345
         Left            =   2400
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   1950
         _ExtentX        =   2990
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Object.Height          =   345
         Object.Width           =   1950
         cSize           =   1
         cChkResult      =   0   'False
         cTypeIntval     =   0
         BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   915
         Index           =   2
         Left            =   2325
         Top             =   -60
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "仕入先"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
         Begin CisText_V60.CisText iB1_SRTorcd 
            Height          =   375
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   420
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXXXX"
            MaxLength       =   6
            cDataReplace    =   1
            cFaZero         =   0
            cFbComma        =   0
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   915
         Index           =   53
         Left            =   3720
         Top             =   -60
         Visible         =   0   'False
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "区"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText iB1_Ku 
            Height          =   375
            Left            =   0
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   420
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
            cFbComma        =   0
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   915
         Index           =   0
         Left            =   3360
         Top             =   -60
         Visible         =   0   'False
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "行"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText iB1_Gyo 
            Height          =   375
            Left            =   90
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   420
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
            cFbComma        =   0
         End
      End
   End
   Begin VB.PictureBox B1_Area2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3645
      Left            =   2228
      ScaleHeight     =   3585
      ScaleWidth      =   9570
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5865
      Width           =   9630
      Begin VB.PictureBox B1_Area9 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H80000004&
         BorderStyle     =   0  'なし
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3465
         Left            =   120
         ScaleHeight     =   3465
         ScaleWidth      =   9300
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   60
         Width           =   9300
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   435
            Index           =   47
            Left            =   9195
            Top             =   90
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   767
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   435
            Index           =   14
            Left            =   3300
            Top             =   90
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   767
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "品     番"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   10
            Left            =   6780
            Top             =   540
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   375
               Index           =   0
               Left            =   240
               TabIndex        =   21
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D B12_Syuyo_Back 
            Height          =   555
            Index           =   0
            Left            =   5505
            Top             =   540
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   375
               Index           =   0
               Left            =   270
               TabIndex        =   20
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   435
            Index           =   18
            Left            =   5505
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   767
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "収容数"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   435
            Index           =   17
            Left            =   6780
            Top             =   90
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   767
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "箱数"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   435
            Index           =   15
            Left            =   2325
            Top             =   90
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   767
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "背番号"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   48
            Left            =   2325
            Top             =   540
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   375
               Index           =   0
               Left            =   210
               TabIndex        =   18
               Top             =   90
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   29
            Left            =   3300
            Top             =   540
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   375
               Index           =   0
               Left            =   150
               TabIndex        =   19
               Top             =   90
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXXXXXXX"
               MaxLength       =   14
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   435
            Index           =   20
            Left            =   7815
            Top             =   90
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   767
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "納入数"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   66
            Left            =   7815
            Top             =   540
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Su 
               Height          =   375
               Index           =   0
               Left            =   270
               TabIndex        =   22
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "######"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   435
            Index           =   6
            Left            =   90
            Top             =   90
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   767
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "№"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   525
            Index           =   59
            Left            =   90
            Top             =   555
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   926
            BackColor       =   14737632
            Caption         =   "１"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   7
            Left            =   6780
            Top             =   1110
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   375
               Index           =   1
               Left            =   240
               TabIndex        =   28
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D B12_Syuyo_Back 
            Height          =   555
            Index           =   1
            Left            =   5505
            Top             =   1110
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   375
               Index           =   1
               Left            =   270
               TabIndex        =   27
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   8
            Left            =   2325
            Top             =   1110
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   375
               Index           =   1
               Left            =   210
               TabIndex        =   25
               Top             =   90
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   9
            Left            =   3300
            Top             =   1110
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   375
               Index           =   1
               Left            =   150
               TabIndex        =   26
               Top             =   90
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXXXXXXX"
               MaxLength       =   14
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   11
            Left            =   7815
            Top             =   1110
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Su 
               Height          =   375
               Index           =   1
               Left            =   270
               TabIndex        =   29
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "######"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   525
            Index           =   13
            Left            =   90
            Top             =   1125
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   926
            BackColor       =   14737632
            Caption         =   "２"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   21
            Left            =   6780
            Top             =   1680
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   375
               Index           =   2
               Left            =   240
               TabIndex        =   35
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D B12_Syuyo_Back 
            Height          =   555
            Index           =   2
            Left            =   5505
            Top             =   1680
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   375
               Index           =   2
               Left            =   270
               TabIndex        =   34
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   22
            Left            =   2325
            Top             =   1680
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   375
               Index           =   2
               Left            =   210
               TabIndex        =   32
               Top             =   90
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   23
            Left            =   3300
            Top             =   1680
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   375
               Index           =   2
               Left            =   150
               TabIndex        =   33
               Top             =   90
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXXXXXXX"
               MaxLength       =   14
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   24
            Left            =   7815
            Top             =   1680
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Su 
               Height          =   375
               Index           =   2
               Left            =   270
               TabIndex        =   36
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "######"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   525
            Index           =   25
            Left            =   90
            Top             =   1695
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   926
            BackColor       =   14737632
            Caption         =   "３"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   26
            Left            =   6780
            Top             =   2250
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   375
               Index           =   3
               Left            =   240
               TabIndex        =   42
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D B12_Syuyo_Back 
            Height          =   555
            Index           =   3
            Left            =   5505
            Top             =   2250
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   375
               Index           =   3
               Left            =   270
               TabIndex        =   41
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   27
            Left            =   2325
            Top             =   2250
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   375
               Index           =   3
               Left            =   210
               TabIndex        =   39
               Top             =   90
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   28
            Left            =   3300
            Top             =   2250
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   375
               Index           =   3
               Left            =   150
               TabIndex        =   40
               Top             =   90
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXXXXXXX"
               MaxLength       =   14
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   30
            Left            =   7815
            Top             =   2250
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Su 
               Height          =   375
               Index           =   3
               Left            =   270
               TabIndex        =   43
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "######"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   525
            Index           =   31
            Left            =   90
            Top             =   2265
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   926
            BackColor       =   14737632
            Caption         =   "４"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   32
            Left            =   6780
            Top             =   2820
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   375
               Index           =   4
               Left            =   240
               TabIndex        =   49
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D B12_Syuyo_Back 
            Height          =   555
            Index           =   4
            Left            =   5505
            Top             =   2820
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   375
               Index           =   4
               Left            =   270
               TabIndex        =   48
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   33
            Left            =   2325
            Top             =   2820
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   375
               Index           =   4
               Left            =   240
               TabIndex        =   46
               Top             =   90
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   34
            Left            =   3300
            Top             =   2820
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   375
               Index           =   4
               Left            =   150
               TabIndex        =   47
               Top             =   90
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXXXXXXX"
               MaxLength       =   14
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   35
            Left            =   7815
            Top             =   2820
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Su 
               Height          =   375
               Index           =   4
               Left            =   270
               TabIndex        =   50
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "######"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   525
            Index           =   36
            Left            =   90
            Top             =   2835
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   926
            BackColor       =   14737632
            Caption         =   "５"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   435
            Index           =   12
            Left            =   1815
            Top             =   90
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   767
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "分"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   435
            Index           =   16
            Left            =   525
            Top             =   90
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   767
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "納　番"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   19
            Left            =   1815
            Top             =   540
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Bunno 
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   17
               Top             =   90
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   37
            Left            =   1815
            Top             =   1110
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Bunno 
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   24
               Top             =   90
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   38
            Left            =   1815
            Top             =   1680
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Bunno 
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   31
               Top             =   90
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   39
            Left            =   1815
            Top             =   2250
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Bunno 
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   38
               Top             =   90
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   40
            Left            =   1815
            Top             =   2820
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Bunno 
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   45
               Top             =   90
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   41
            Left            =   525
            Top             =   540
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Noubn 
               Height          =   375
               Index           =   0
               Left            =   300
               TabIndex        =   16
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
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
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   42
            Left            =   525
            Top             =   1110
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Noubn 
               Height          =   375
               Index           =   1
               Left            =   300
               TabIndex        =   23
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
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
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   43
            Left            =   525
            Top             =   1680
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Noubn 
               Height          =   375
               Index           =   2
               Left            =   300
               TabIndex        =   30
               Top             =   120
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
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
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   44
            Left            =   525
            Top             =   2250
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Noubn 
               Height          =   375
               Index           =   3
               Left            =   300
               TabIndex        =   37
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
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
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   45
            Left            =   525
            Top             =   2820
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            Begin CisText_V60.CisText iB12_Noubn 
               Height          =   375
               Index           =   4
               Left            =   300
               TabIndex        =   44
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
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
         End
      End
   End
   Begin VB.Timer Timer2 
      Left            =   2265
      Top             =   390
   End
   Begin VB.Timer Timer1 
      Left            =   2700
      Top             =   390
   End
   Begin VB.TextBox Dummy 
      Height          =   270
      Left            =   1965
      TabIndex        =   82
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   480
      Width           =   180
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
      Caption         =   "【 受注チケット入力（アキタ）  】"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
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
         BorderStyle     =   0  'なし
         Height          =   345
         Index           =   1
         Left            =   10140
         ScaleHeight     =   345
         ScaleWidth      =   3615
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   30
         Width           =   3615
         Begin Cis3D_v60.CIS3D UC_3D6 
            Height          =   345
            Left            =   30
            Top             =   0
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   609
            BackColor       =   8388736
            ForeColor       =   16777215
            Caption         =   "読取件数"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cFont3D         =   2
            cAlingnment     =   1
            cPositionX      =   90
            Begin Cis3D_v60.CIS3D H1lb_OCRCnt 
               Height          =   255
               Left            =   1110
               Top             =   60
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   8388736
               ForeColor       =   16777215
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
               cFont3D         =   2
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D7 
            Height          =   345
            Left            =   1830
            Top             =   0
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   609
            BackColor       =   8454143
            ForeColor       =   0
            Caption         =   "手入力件数"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
               Left            =   1200
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
      End
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Index           =   1
         Left            =   14220
         Top             =   90
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
         Top             =   60
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   503
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXXXXXXX2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
   Begin CisText_V60.CisText iB14_Yobi1 
      Height          =   345
      Left            =   9225
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   609
      cFormat         =   "#"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "9"
      MaxLength       =   1
      cDataType       =   1
      cDataReplace    =   1
      cFbComma        =   0
      cGFormat        =   "#"
      cILength        =   1
   End
   Begin CisText_V60.CisText iB14_Bun 
      Height          =   345
      Left            =   10575
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "X"
      MaxLength       =   1
      cDataReplace    =   1
      cFbComma        =   0
   End
   Begin CisText_V60.CisText iB14_HBin 
      Height          =   345
      Left            =   10965
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   4350
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   609
      cFormat         =   "00"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "99"
      MaxLength       =   2
      cDataType       =   1
      cDataReplace    =   1
      cFaZero         =   0
      cFdAutoFormat   =   1
      cGFormat        =   "##"
      cILength        =   2
   End
   Begin CisText_V60.CisText iB14_MotoYobi 
      Height          =   345
      Left            =   11205
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   609
      cFormat         =   "####"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "9999"
      MaxLength       =   4
      cDataType       =   1
      cDataReplace    =   1
      cFbComma        =   0
      cGFormat        =   "####"
      cILength        =   4
   End
   Begin CisText_V60.CisText iB14_Kouku 
      Height          =   345
      Left            =   10845
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
      cFormat         =   "##"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "99"
      MaxLength       =   2
      cDataType       =   1
      cDataReplace    =   1
      cFbComma        =   0
      cGFormat        =   "##"
      cILength        =   2
   End
   Begin CisText_V60.CisText iB14_HKB 
      Height          =   345
      Left            =   10305
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   609
      cFormat         =   "#"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "9"
      MaxLength       =   1
      cDataType       =   1
      cDataReplace    =   1
      cFbComma        =   0
      cGFormat        =   "#"
      cILength        =   1
   End
   Begin CisText_V60.CisText iB14_Gyo 
      Height          =   345
      Left            =   10035
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   609
      cFormat         =   "#"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "9"
      MaxLength       =   1
      cDataType       =   1
      cDataReplace    =   1
      cFbComma        =   0
      cGFormat        =   "#"
      cILength        =   1
   End
   Begin CisText_V60.CisText iB14_Yobi3 
      Height          =   345
      Left            =   9765
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   609
      cFormat         =   "#"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "9"
      MaxLength       =   1
      cDataType       =   1
      cDataReplace    =   1
      cFbComma        =   0
      cGFormat        =   "#"
      cILength        =   1
   End
   Begin CisText_V60.CisText iB14_Yobi2 
      Height          =   345
      Left            =   9495
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   609
      cFormat         =   "#"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "9"
      MaxLength       =   1
      cDataType       =   1
      cDataReplace    =   1
      cFbComma        =   0
      cGFormat        =   "#"
      cILength        =   1
   End
   Begin CisYMD_V60.CisYMD iB14_HYmd 
      Height          =   330
      Left            =   9225
      TabIndex        =   61
      Top             =   4350
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3069
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Object.Width           =   1755
      cChkResult      =   0   'False
   End
   Begin Cis3D_v60.CIS3D K_Sykbnm 
      Height          =   795
      Left            =   2228
      Top             =   900
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   1402
      Caption         =   "読取"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   20.25
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
   Begin Cis3D_v60.CIS3D H1_Area1 
      Height          =   1095
      Left            =   2265
      Top             =   2955
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   1931
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
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
      Begin Cis3D_v60.CIS3D CIS3D17 
         Height          =   915
         Left            =   90
         Top             =   90
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "受注区分"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "指示"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   2100
            TabIndex        =   1
            Top             =   120
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "特配"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   2115
            TabIndex        =   5
            Top             =   525
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "臨時"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   1185
            TabIndex        =   4
            Top             =   525
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H008080FF&
            Caption         =   "通常"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1170
            TabIndex        =   0
            Top             =   120
            Value           =   -1  'True
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "号試"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   2985
            TabIndex        =   2
            Top             =   120
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "その他"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   3930
            TabIndex        =   7
            Top             =   525
            Width           =   1320
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "スポット"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   3915
            TabIndex        =   3
            Top             =   120
            Width           =   1320
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "代替"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   3000
            TabIndex        =   6
            Top             =   525
            Width           =   840
         End
      End
   End
   Begin VB.Label ReadColor 
      BackColor       =   &H00FFC0C0&
      Caption         =   "読取中色"
      Height          =   255
      Left            =   90
      TabIndex        =   64
      Top             =   1020
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "CXJ0550"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  受注チケット入力（アキタ By 土岐津Ver）
'**       フォームID    :  CXJ0550
'**       処理概要      :
'**
'**       作  成  日    :  2010/03/03  By CIS
'**       変  更  日    :  2010/04/19  By CIS  処理後の画面ｸﾘｱｰ中止
'**       変  更  日    :  2010/11/04  By CIS  受注区分判断修正
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim SyoriKb         As String       ' 処理区分(追加･修正･削除)
    Dim SyoriNM         As String       ' 処理名称(追加･修正･削除)
    Dim ErrCnt          As Integer
    Dim Cnt_OCR         As Integer      ' OCR処理件数
    Dim Cnt_ADD         As Integer      ' 追加処理件数
    Dim wEndChk         As Boolean
    
    Dim wObj            As Object
    Dim wNum1           As Integer
    Dim wNum2           As Integer
    Dim wNumL           As Long
    Dim wStr1           As String
    Dim wStr2           As String
    Dim wStr3           As String
    Dim wIndex          As Integer
    Dim wI              As Integer
    Dim w_JKbn          As Byte         '受注区分
    Dim sFlg            As Integer

    Dim mERMsg          As Boolean
    Dim mFocusObj       As Object
'   * 伝票内容退避ワーク *
    Dim wcc_Head        As String
    Dim wcc_Body        As String
    Dim wGyo            As Integer
'   * 対象伝票№
    Dim wDenOK          As String       ' 2009.11.02 add
    Dim wAkitaCode      As String
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   # 二 重 起 動 防 止 #
'   #------------------#
    If Not Dupli_Start Then End
'   #------------------#
'   # 初 期 内 容 取 得 #
'   #------------------#
    If Not IniGet Then End
'   #------------------#
'   #  項 目 初 期 化   #
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

    Timer1.Interval = 1000  ' 時刻表示間隔
    Timer1.Enabled = True   ' 時刻表示開始
    Timer2.Interval = 500   ' 処理区分点滅間隔
    Timer2.Enabled = True   ' 処理区分点滅禁止
'
'    wDenOK = "4762"         ' 対象伝票番号  ' 2009.11.02
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    Call QT10QSRInfoGet
    ProcHB = "H1"
    If Not QT10QSR_UseFlg Then
       SyoriKb = "ADD"
    Else
       SyoriKb = "RED"
    End If
    Call DispChange(ProcHB)
    Call SyoriName(SyoriKb)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)
    
'   >> 名称マスタ確認
    Call MeisyoInit
    
    H1Op_JKbn(1).Value = True
    Call H1Op_JKbn_Click(1)

    
'>> アキタコード作成
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "アキタ伝票"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "アキタコード"
        .作成者 = gTanto
        .検索名 = ""
        .コード名 = ""
        .名称名 = ""
        .英数字1 = "1048"
        If Not SYMInitCreate(True) Then
        End If
    End With
    
    
    If MeisyoGet("アキタ伝票", 1) Then
        wAkitaCode = SYM.英数字1
    Else
        wAkitaCode = "1048"
    End If
    
    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    Cnt_OCR = 0: Cnt_ADD = 0
   
    If InitProc Then        ' 初期処理（データ削除）
        wEndChk = True
    Else
        wEndChk = False
        Call PB_END_Click
    End If

'    iH1_YMD.cInitDate = True
    
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》       +
'+----------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    If Not FormAct Then Exit Function
    If UCase(IVENT) <> "DOWN" Then Exit Function
    
    Key_Acc = False
    
    Select Case Key_Code
        Case vbKeyF1:       If PB_ADD.Visible Then Call PB_Add_Click    '【新規】
        Case vbKeyF5:       If PB_RED.Visible Then Call PB_RED_Click    '【読取】
        Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
        Case vbKeyF10:      If ProcHB = "B2" Then Call Test_OnComm    'テスト用
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
        Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
        Case vbKeyReturn
                            If PB_ENT.Visible Then                      '【入力】
                                Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                Call PB_ENT_Click
                                Dummy.Enabled = False
                            End If
        Case Else:          Exit Function
    End Select
    
    Key_Acc = True
End Function
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
'   #### 終了処理 ####
    Call EndProc
'   #### ﾃﾞｰﾀﾍﾞｰｽ切断 ####
    Call CisDB.DBDISConnect
'   #### ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去 ####
    Call CisFrm.UnLoadDisp
    End
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call ReturnPress
    Dummy.Enabled = False
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call CisFrm.MousePT(11)
    
    If ProcHB = "B1" Then
        If SyoriKb <> "RED" Then
           ProcHB = "H1"
           iH1_Ymd = ""
           iH1_Bin = ""
           Call DispChange(ProcHB)
           Call SyoriName(SyoriKb)
           Call iH1_Ymd.SetFocus        ' 2009.11.02 add
        Else
           ProcHB = "B2"
           Call DispChange(ProcHB)
           Call SyoriName(SyoriKb)
        End If
        Call HeadBodyClear("B")
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B2" Then
        ProcHB = "H1"
        iH1_Ymd = ""                    ' 2009.11.05 add
        iH1_Bin = ""                    ' 2009.11.05 add
        Call HeadBodyClear("B")
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKb)
        Call iH1_Ymd.SetFocus           ' 2009.11.05 add
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    新  規  キ ー(F1)       +
'+----------------------------+
Private Sub PB_Add_Click()
    SyoriKb = "ADD"
    Call SyoriName(SyoriKb)
    If ProcHB = "B2" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
    End If
    Call iH1_Ymd.SetFocus       ' 2009.11.02 add
End Sub
'+----------------------------+
'+      読取 キ ー(F5)         +
'+----------------------------+
Private Sub PB_RED_Click()
    SyoriKb = "RED"
    Call SyoriName(SyoriKb)
    If ProcHB = "B2" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
    End If
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Ymd" Then iH1_Ymd.ShowCalender: Exit Sub
'    If PB_Look.Tag = "iB13_MM" Then iB1_MM.ShowCalender: Exit Sub
'    If PB_Look.Tag = "iB13_DD" Then iB1_DD.ShowCalender: Exit Sub
    If PB_Look.Tag Like "iB12_Hin*" Then Call RefHinbn(Right(PB_Look.Tag, 1))
    If PB_Look.Tag Like "iB12_Seban*" Then Call RefHinbn(Right(PB_Look.Tag, 1))
End Sub
'+------------------------------------+'
'+                                    +
'+      参               照           +
'+                                    +
'+------------------------------------+'
'+-------------------------------+
'+    品  番  検  索              +
'+-------------------------------+
Private Sub RefHinbn(Index As Integer)

    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 0
    RVI_SK = 9
    RVI_Torcd = ""
    RVI_Ukeir = ""
    RVI_Seban = ""
    RVI_Tehai = ""

    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        If Mid(RV_Hinbn, 6, 5) <> "" Then
           If Len(RV_Hinbn) > 10 Then                                               ' 2009.11.05 add
              iB12_Hinbn(Index) = Mid(RV_Hinbn, 1, 5) & "-" & Mid(RV_Hinbn, 6, 5) & "-" & Mid(RV_Hinbn, 11, 2)
           Else                                                                     ' 2009.11.05 add
              iB12_Hinbn(Index) = Mid(RV_Hinbn, 1, 5) & "-" & Mid(RV_Hinbn, 6, 5)   ' 2009.11.05 add
           End If                                                                   ' 2009.11.05 add
        Else
'           iB12_Hinbn(Index) = Mid(RV_Hinbn, 1, 5) & "-"                           ' 2009.11.05 del
        End If
        If Trim(RV_Seban) <> "" Then
           iB12_Seban(Index) = Trim(RV_Seban)
        End If
    End If

End Sub
'************************************'
'*        画面フィールド
'************************************'
'*---------------------------------------------------*
'*      iH1_YMD   ( GotFocus )        納入日
'*---------------------------------------------------*
Private Sub iH1_YMD_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
'*---------------------------------------------------*
'*      iH1_YMD   ( LostFocus )       納入日
'*---------------------------------------------------*
Private Sub iH1_YMD_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Visible = False
        PB_Look.Tag = ""
    End If
End Sub
'*---------------------------------------------------*
'*      iB12_Hinbn   ( GotFocus )        品番
'*---------------------------------------------------*
Private Sub iB12_Hinbn_GotFocus(Index As Integer)
    PB_Look.Tag = ActiveControl.Name & Index
    PB_Look.Visible = True
End Sub
'*---------------------------------------------------*
'*      iB12_Hinbn   ( LostFocus )       品番
'*---------------------------------------------------*
Private Sub iB12_Hinbn_LostFocus(Index As Integer)
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Visible = False
        PB_Look.Tag = ""
    End If
End Sub
'*---------------------------------------------------*
'*      iB12_Seban   ( GotFocus )        背番号
'*---------------------------------------------------*
Private Sub iB12_Seban_GotFocus(Index As Integer)
    PB_Look.Tag = ActiveControl.Name & Index
    PB_Look.Visible = True
End Sub
'*---------------------------------------------------*
'*      iB12_Seban   ( LostFocus )       背番号
'*---------------------------------------------------*
Private Sub iB12_Seban_LostFocus(Index As Integer)
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Visible = False
        PB_Look.Tag = ""
    End If
End Sub
'*---------------------------------------------------*
'*      iB12_Hako   ( LostFocus )       箱数
'*---------------------------------------------------*
Private Sub iB12_Hako_LostFocus(Index As Integer)
    If RTrim(iB12_Hinbn(Index)) = "" Then Exit Sub
'   * 省略時 １ *
    If CisFun.Val2(iB12_Hako(Index)) = 0 Then
       If CisFun.Val2(iB12_Syuyo(Index)) = 0 Then
          iB12_Hako(Index) = "  1"
       Else
          gInt = CisFun.Val2(iB12_Su(Index)) \ CisFun.Val2(iB12_Syuyo(Index))
          gLong = CisFun.Val2(iB12_Su(Index)) Mod CisFun.Val2(iB12_Syuyo(Index))
          If gLong <> 0 Then
             gInt = gInt + 1
          End If
          iB12_Hako(Index) = CisFun.RSetFld(gInt, 3, "###")
       End If
    End If
'   * 納入数計算 *
    If CisFun.Val2(iB12_Su(Index)) <> 0 Then Exit Sub
    wNumL = CisFun.Val2(iB12_Hako(Index)) * CisFun.Val2(iB12_Syuyo(Index))
    iB12_Su(Index) = Format(Format(wNumL, "#"), String(6, "@"))
End Sub
'*---------------------------------------------------*
'*      iB12_Syuyo  ( LostFocus )       収容数
'*---------------------------------------------------*
Private Sub iB12_Syuyo_LostFocus(Index As Integer)
    If RTrim(iB12_Hinbn(Index)) = "" Then Exit Sub
'   * 省略時 １ *
    If CisFun.Val2(iB12_Syuyo(Index)) = 0 Then iB12_Syuyo(Index) = Format(Format(CisFun.Val2(iB12_Su(Index)), "#"), String(5, "@"))

'   * 納入数計算 *
    If CisFun.Val2(iB12_Su(Index)) <> 0 Then Exit Sub
    wNumL = CisFun.Val2(iB12_Hako(Index)) * CisFun.Val2(iB12_Syuyo(Index))
    iB12_Su(Index) = Format(Format(wNumL, "#"), String(6, "@"))
End Sub
'*---------------------------------------------------*
'*      iB12_Su     ( LostFocus )       納入数
'*---------------------------------------------------*
Private Sub iB12_Su_LostFocus(Index As Integer)
    If CisFun.Val2(iB12_Su(Index)) <> 0 Then Exit Sub

    wNumL = CisFun.Val2(iB12_Hako(Index)) * CisFun.Val2(iB12_Syuyo(Index))
    iB12_Su(Index) = Format(Format(wNumL, "#"), String(6, "@"))
End Sub
'+----------------------------------------+
'+  H1Op_Jkbn ( Click )      受注区分
'+----------------------------------------+
Private Sub H1Op_JKbn_Click(Index As Integer)
'    H1Op_JKbn(1).Tag = Index           2010/11/04
    For wIndex = 1 To 8
       H1Op_JKbn(wIndex).BackColor = &HC0C0C0
    Next wIndex
    H1Op_JKbn(Index).BackColor = &H8080FF
    Select Case Index
        Case 1:     Back_Ymd.Caption = "納 入 日"
        Case Else:  Back_Ymd.Caption = "納 入 日"
    End Select
End Sub
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   
    Dim W_Color     As Long
    
    W_Color = &H80000008 '初期色ｾｯﾄ(黒)
    K_Sykbnm.BackColor = &HC0C0C0
    Select Case SyoriKb
     Case "RED"
          SyoriNM = "QR読取"
          K_Sykbnm.ForeColor = W_Color   '(黒)
          PB_ADD.Visible = True
          PB_RED.Visible = False
     Case "ADD"
          SyoriNM = "手入力"
          K_Sykbnm.ForeColor = &HFFFF&   '(黄色)
          PB_ADD.Visible = False
          PB_RED.Visible = True
     Case Else
          SyoriNM = ""
    End Select
    
    Timer2.Enabled = False
    If SyoriNM = "" Then
        K_Sykbnm.Visible = False
        K_Sykbnm.Caption = ""
        Exit Sub
    End If
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    Timer2.Enabled = True
    
    If Not QT10QSR_UseFlg Then
       PB_ADD.Visible = False
       PB_RED.Visible = False
    End If
    
End Sub
'****************************
'*                          *
'*      HEAD/BODY 処理      *
'*                          *
'****************************
'****************************
'*      入力切替 処理        *
'****************************
Private Function DispChange(Pro As String)
    Dim H1Mode      As Boolean
    Dim H1Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long
    Dim B2Mode      As Boolean
    Dim B2Color     As Long

    H1Mode = False: H1Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
    B2Mode = False: B2Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1":   H1Mode = True:  H1Color = gPGotSel
       Case "B1":   B1Mode = True:  B1Color = gPGotSel
       Case "B2":   B2Mode = True:  B2Color = gPGotSel
    End Select
        
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    H1_Area2.Enabled = H1Mode
    H1_Area2.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B1_Area2.Enabled = B1Mode
    B1_Area2.BackColor = B1Color

    PB_CAN.Visible = Not H1Mode

    If ProcHB = "B2" Then
        SyoriNM = "【 読 取 中 】"
        K_Sykbnm.BackColor = ReadColor.BackColor
        Set QT10QSR_Comm = PDQComm1
        Call QT10QSRInfoGet
        If Not QT10QSR_Open Then
           End
        End If
    End If
    
    PB_ADD.Visible = H1Mode
    PB_RED.Visible = H1Mode
    
    PB_ENT.Visible = Not B2Mode
    
    If Not QT10QSR_UseFlg Then
       PB_ADD.Visible = False
       PB_RED.Visible = False
    End If
' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'**************************************
'*      Head/Body 初期化処理           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String

    If ClrType = "ALL" Then wbc_Enm = HeadBody & "*" Else wbc_Enm = Trim(ClrType)
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then .text = "" ': .Enabled = True
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Cmb_*" Then .Clear
        End With
    Next gObj
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
'   * 納入日 *
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        If SyoriKb = "RED" Then
            ProcHB = "B2"
            Call DispChange(ProcHB)
            Call ChiketoProc                         ' QR読取処理
            GoTo ReturnPress_Ed
        End If
        ProcHB = "B1"
        Call DispChange(ProcHB)
        iB1_Kibn = "F8"             ' 2009.11.01 add
        GoTo ReturnPress_Ed
    End If

'   * 伝票入力 *
    If ProcHB = "B1" Then
        Call Line_Cut
        If Not ContentsChk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        If SyoriKb = "ADD" Then iB1_Kibn = "F8"     ' 2009.11.05 add : 機番初期値 = F8
        
        If SyoriKb = "ADD" Then ProcHB = "B1" Else ProcHB = "B2"
        
        Call DispChange(ProcHB)
        
        GoTo ReturnPress_Ed
    End If

ReturnPress_Ed:
    Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理(検収日・社内区分)
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
 
    If Trim(iH1_Ymd) <> "" Then
        If Not iH1_Ymd.cChkResult Then
            With CisFun
                .MB_Lines = 5
                .MB_Button = Error
                .MB_Title = ""
                .MB_MSG(2) = "    納入日を正しく入力して下さい        "
                .MBOX
            End With
            Call iH1_Ymd.SetFocus       ' 2009.11.01 add
            Exit Function
        End If
    End If
    If SyoriKb = "ADD" Then
        If RTrim(iH1_Ymd) = "" Then
            With CisFun
                .MB_Lines = 5
                .MB_Button = Error
                .MB_Title = ""
                .MB_MSG(2) = "    納入日を入力して下さい        "
                .MBOX
            End With
            Call iH1_Ymd.SetFocus       ' 2009.11.01 add
            Exit Function
        End If
    End If
'
    For wIndex = 1 To 8
        If H1Op_JKbn(wIndex).Value Then
            w_JKbn = CisFun.Val2(H1Op_JKbn(wIndex).Tag)
            Exit For
        End If
    Next wIndex
    
    Head1Chk = True
End Function
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   Call CisFrm.TimeSet(1)
End Sub
'*******************************
'*      処理区分名点滅制御      *
'*******************************
Private Sub Timer2_Timer()
   If K_Sykbnm.Caption <> "" Then
      K_Sykbnm.Caption = ""
   Else
      K_Sykbnm.Caption = SyoriNM
   End If
End Sub
'**************************************
'*                                    *
'*      かんばんチケット読取処理      *
'*                                    *
'**************************************
Private Function ChiketoProc()
    
    ChiketoProc = False
' 232C ｵﾌﾞｼﾞｪｸﾄ初期設定
    Set QT10QSR_Comm = PDQComm1
    Call QT10QSRInfoGet
    If Not QT10QSR_Open Then
        Call PB_CAN_Click
    End If
    
    On Error GoTo 0
    Exit Function

OcrProc_Err:
    Call CisFun.ErrorBox("ﾁｹｯﾄ QR読取")
    End
    
End Function
'+------------------------------+
'+  読取処理     *
'+------------------------------+
Private Sub PDQComm1_OnComm()
    Static wPNo     As Integer
    Static Rtn      As String
    Dim wComm       As String
    
    gStr = QT10QSR_Event
    If gStr = "ERR" Then Exit Sub
    If gStr <> "RCV" Then Exit Sub
    
    If Not QT10QSR_RCVProc Then Exit Sub    ' CR が存在しない時抜ける
''-----------------------------------------------------------------------' 2009.10.30 add start
'    If Mid(QT10QSR_ReadData, 14, 4) <> wDenOK Then
'        With CisFun
'             .MB_Lines = 5
'             .MB_MSG(2) = "  対象外伝票です      "
'             .MB_Title = ""
'             .MB_Button = Error
'             .MBOX
'        End With
'        QT10QSR_ReadData = ""
'        Exit Sub
'    End If
''-----------------------------------------------------------------------' 2009.10.30 add end

    Call ReadDsp(QT10QSR_ReadData)          '読取内容画面表示
    DoEvents
    Call Line_Cut
    
'    If Not ContentsChk Then Exit Sub
'    wcc_Head = Mid(QT10QSR_ReadData, 1, 56)
'    wcc_Body = Mid(QT10QSR_ReadData, 57)
    
'
'    If ChekData(0, 0, 1) Then
    


''''Call QRConvert(QT10QSR_ReadData)    ' PXJ0610の例
''''QT10QSR_ReadData = ""               ' PXJ0610の例
''''If Not QRProc Then Exit Sub         ' PXJ0610の例
    
    
    If ContentsChk Then
       H1lb_OCRCnt = Format(CisFun.Val2(H1lb_OCRCnt) + 1, "#,###")
'''       Call HeadBodyClear("B")                                               2010/04/19 処理後クリアー中止
       DoEvents
    Else
       Call QT10QSR_Close
       ProcHB = "B1"
       Call SyoriName(SyoriKb)
       Call DispChange(ProcHB)
    
       On Error Resume Next
       mFocusObj.SetFocus
       Set mFocusObj = Nothing
       On Error GoTo 0
    End If
    
    QT10QSR_ReadData = ""
End Sub
'+------------------------------------------+
'+
'+------------------------------------------+
Private Function ChekData(ReadType As Integer, OutType As Integer, ChekFlg As Integer) As Boolean

    Dim wErNo           As Integer
    Dim wErGyo          As Integer
    
    ChekData = False
    
    mERMsg = False

    With CisDB
        .SQL = "受注アキタ納品書チェック"
        .StoadoCount = 12
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用      ' 処理区分 0:ﾘｰﾀﾞｰ,1:手入力
        .ParaNo = 2: .ParaIO = Input用      ' 受注区分
        .ParaNo = 3: .ParaIO = Input用      ' 納入日
        .ParaNo = 4: .ParaIO = Input用      ' 納入便
        .ParaNo = 5: .ParaIO = Input用      ' 納入先～受注者事業所・納品書№・ページ
        .ParaNo = 6: .ParaIO = InOut用      ' 伝票の明細部(品番～納入数<5明細分>)
        .ParaNo = 7: .ParaIO = InOut用      ' 伝票の明細部(品番～納入数<5明細分>)
        .ParaNo = 8: .ParaIO = Input用      ' アキタ自社コード
        .ParaNo = 9: .ParaIO = OutPut用     ' 変換後 受入
        .ParaNo = 10: .ParaIO = OutPut用     ' 変換後 納入先
        .ParaNo = 11: .ParaIO = OutPut用    ' チェックエラーNO
        .ParaNo = 12: .ParaIO = OutPut用    ' チェックエラー行NO

        .ParaNo = 1: .ParaValue = ReadType  ' 0:ﾘｰﾀﾞｰ,1:手入力
        '.ParaNo = 2: .ParaValue = H1Op_JKbn(1).Tag             '2010/11/04
        .ParaNo = 2: .ParaValue = w_JKbn                        '2010/11/04
        .ParaNo = 3: .ParaValue = iH1_Ymd
        .ParaNo = 4: .ParaValue = iH1_Bin
        .ParaNo = 5: .ParaValue = Mid(wcc_Head, 1, 55)
        .ParaNo = 6: .ParaValue = wcc_Body
        .ParaNo = 7: .ParaValue = wcc_Body ''2
        .ParaNo = 8: .ParaValue = wAkitaCode

        If Not .DBStored Then End
        ' エラー番号
        .ParaNo = 11: wErNo = CisFun.Val2(.ParaValue)
        ' エラー行No
        .ParaNo = 12: wErGyo = CisFun.Val2(.ParaValue)
        
    End With
'--              1 : 対象外チケット           2 : 発注区分エラー
'--              3 : 納入先変換エラー         4 : 納入先未登録
'--              5 : 自社向伝票でない         6 :
'--              7 : 収容数違い               8 : 納入数エラー(箱数×収容数≠納入数)
'--              9 : 伝票違い
'--              15: 品番取引先マスタ未登録
'--              16: 受注重複エラー(受注テーブル)
'
'--              81: 受注管理NO獲得エラー
'--              82: 出荷指示書NO獲得エラー
'--              85: 受注データ出力エラー
'--              86: 出荷データ出力エラー
'--              99: データ更新エラー
    
    With CisFun
        .MB_Lines = 5
        .MB_Title = ""
        .MB_Button = Error
            
'        If wErNo = 8 And ChekFlg = 0 Then
'            .MB_MSG(1) = "【  " & StrConv(wErGyo, vbWide) & " 行目  】"
'            .MB_MSG(3) = "   納入数が異なる ( 箱数×収容数 )    "
'            Set mFocusObj = iB12_Hako(wErGyo - 1)
'        End If

'        .MB_MSG(2) = "未定義エラー：" & Format(wErNo, "00") & " / 行：" & Format(wErGyo, "0") & "          "

        Select Case wErNo
            Case 1
                    .MB_MSG(2) = "     納入先未登録      "
                    Set mFocusObj = iB1_Hacyu
            Case 4
                    .MB_MSG(2) = "     受入未登録        "
                    Set mFocusObj = iB1_Uke
'-----------------------------------------------------------------------------------------------------------' 2009.10.30 del start
'            Case 5
'                    .MB_MSG(2) = "     処理対象伝票でない(＃チェック)      "
'                    Set mFocusObj = iB1_Sharp
'            Case 6
'                    .MB_MSG(2) = "     チェックデジットエラー：入力内容を確認して下さい（ＨＥＡＤ部）  "
'                    Set mFocusObj = iB1_CD
'-----------------------------------------------------------------------------------------------------------' 2009.10.30 del end
            Case 3
                    .MB_MSG(2) = "     自社向けチケットではありません。      "
                    Set mFocusObj = iB1_SRTorcd
            Case 15
                    .MB_MSG(2) = "【  " & StrConv(wErGyo, vbWide) & " 行目  】"
                    .MB_MSG(4) = "       品番取引先マスタ未登録     "
                    Set mFocusObj = iB12_Hinbn(wErGyo - 1)
            Case 7
                    .MB_MSG(1) = "【  " & StrConv(wErGyo, vbWide) & " 行目  】"
                    .MB_MSG(3) = "   収容数エラー ( 品番取引先マスタ )    "
                    Set mFocusObj = iB12_Syuyo(wErGyo - 1)
            Case 8
                    .MB_MSG(1) = "【  " & StrConv(wErGyo, vbWide) & " 行目  】"
                    .MB_MSG(3) = "   納入数が異なる ( 箱数×収容数 )    "
                    Set mFocusObj = iB12_Hako(wErGyo - 1)
'-----------------------------------------------------------------------------------------------------------' 2009.11.04 rep start
            Case 9
'                   .MB_MSG(1) = "【  " & StrConv(wErGyo, vbWide) & " 行目  】"
                    .MB_MSG(2) = "   処理対象伝票ではありません    "
                    Set mFocusObj = iB1_Hacyu
'-----------------------------------------------------------------------------------------------------------' 2009.11.04 rep end
            Case 16
                    .MB_MSG(2) = "       チケット入力済みです      "
                    Set mFocusObj = iB1_Ku
            Case 17
                    .MB_MSG(2) = "       チケット重複読取です      "
                    Set mFocusObj = iB1_Ku
            Case 88
                    .MB_MSG(2) = "       受注管理NO獲得エラー      "
                    Set mFocusObj = iB1_Ku
            Case 99
                    .MB_MSG(2) = "       受注ワーク出力エラー      "
                    Set mFocusObj = iB1_Ku
'            Case 5
'                    .MB_MSG(2) = "     自社向けチケットではありません。      "
'                    Set mFocusObj = iB1_SRTorcd
'                    If OutType = 0 Then mERMsg = True
'            Case 9
'                    .MB_MSG(2) = "【  " & StrConv(wErGyo, vbWide) & " 行目  】"
'                    .MB_MSG(4) = "       納入数未入力     "
'                    Set mFocusObj = iB12_Su(wErGyo - 1)
'            Case 15
'                    .MB_MSG(2) = "【  " & StrConv(wErGyo, vbWide) & " 行目  】"
'                    .MB_MSG(4) = "       品番マスタ未登録     "
'                    Set mFocusObj = iB12_Hin1(wErGyo - 1)
'            Case 81
'                    .MB_MSG(2) = "       受注管理NO獲得エラー      "
'                    Set mFocusObj = iB13_YMD
'            Case 82
'                    .MB_MSG(2) = "       出荷指示書NO獲得エラー      "
'                    Set mFocusObj = iB13_YMD
'            Case 85
'                    .MB_MSG(2) = "       受注データ出力エラー      "
'                    Set mFocusObj = iB13_YMD
'            Case 86
'                    .MB_MSG(2) = "       出荷データ出力エラー      "
'                    Set mFocusObj = iB13_YMD
        End Select
    End With


'   * 画面再表示 *
    Call HeadBodyClear("B")
' エラー表示後再読み込み
    If mERMsg Then
       Call CisFun.MBOX
       Exit Function
    End If
'   * 画面再表示 *
    iB1_Ku = Mid(wcc_Head, 5, 1)
'    iB1_Sharp = Mid(wcc_Head, 2, 1)
    iB1_SRTorcd = Mid(wcc_Head, 7, 6)
    iB1_Hacyu = Mid(wcc_Head, 14, 5)
    If iB1_Hacyu = "" Then
       iB1_Hacyu = RTrim(wAkitaCode)
    End If
    iB1_Uke = Mid(wcc_Head, 19, 2)
    iB1_Kibn = Mid(wcc_Head, 23, 2)
''    iB1_YY = Mid(wcc_Head, 38, 2)           ' 2009.11.05 add
''    iB1_MM = Mid(wcc_Head, 40, 2)           ' 2009.11.05 upd : 32-->40
''    iB1_DD = Mid(wcc_Head, 42, 2)           ' 2009.11.05 upd : 34-->40
    iB1_Ymd = "20" & Mid(wcc_Head, 38, 6)   ' 2009.11.05 add
    iB1_Bin = Mid(wcc_Head, 44, 2)          ' 2009.11.05 upd : 36-->40
    iB1_Gyo = Mid(wcc_Head, 4, 1)
'    iB1_CD = Mid(wcc_Head, 25, 2)
    
    wStr1 = wcc_Body
'
    For wNum1 = 0 To 4
        wNum2 = wNum1 * 53
        If RTrim(Mid(wStr1, wNum2 + 1, 5)) = "" Then Exit For               ' 納入番号が空白なら終了
'
'        iB12_UkeNo(wNum1) = RTrim(Mid(wStr1, wNum2 + 1, 5))                '受付№
        iB12_Bunno(wNum1) = RTrim(Mid(wStr1, wNum2 + 6, 1))                 '分納
        iB12_Noubn(wNum1) = RTrim(Mid(wStr1, wNum2 + 1, 5))                 '納入番号
        iB12_Seban(wNum1) = RTrim(Mid(wStr1, wNum2 + 7, 4))                 '背番号
        iB12_Hinbn(wNum1) = RTrim(Mid(wStr1, wNum2 + 11, 13))               '品番
        iB12_Syuyo(wNum1) = CisFun.RSetFld(Mid(wStr1, wNum2 + 33, 5), 0, "#####")
        iB12_Hako(wNum1) = CisFun.RSetFld(Mid(wStr1, wNum2 + 38, 3), 0, "###")
        iB12_KiriF(wNum1) = CisFun.RSetFld(Mid(wStr1, wNum2 + 47, 1), 0, "#")
        iB12_KiriZ(wNum1) = CisFun.RSetFld(Mid(wStr1, wNum2 + 48, 6), 0, "######")
'        iB12_CD(wNum1) = RTrim(Mid(wStr1, wNum2 + 44, 2))                  'CD
'
        If CisFun.Val2(Mid(wStr1, wNum2 + 41, 6)) <> 0 Then
            iB12_Su(wNum1) = CisFun.RSetFld(RTrim(Mid(wStr1, wNum2 + 41, 6)), 0, "######")
        End If
    Next wNum1
    
    If wErNo = 0 Then
        ChekData = True
        Exit Function
    End If
    
    If wErNo = 16 Then
       mERMsg = True
    End If
'
    Call CisFun.MBOX
    
    If wErNo > 50 Then
        Call PB_END_Click
    End If
    
    SyoriKb = "ADD"
    Call SyoriName(SyoriKb)
    ProcHB = "B1"
    Call DispChange(ProcHB)

End Function
'**************************************
'*                                    *
'*          その他ファンクション      *
'*                                    *
'**************************************
'+----------------------------------------------------------------------+
'+          明 細 行  行 詰 め
'+----------------------------------------------------------------------+
Private Sub Line_Cut()
    Static wlc_Flg      As Boolean
    
    For wNum1 = 0 To 4
''''    If RTrim(iB12_UkeNo(wNum1)) = "" And RTrim(iB12_Bunno(wNum1)) = "" And RTrim(iB12_Noubn(wNum1)) = "" And RTrim(iB12_Hinbn(wNum1)) = "" And _
           RTrim(iB12_Seban(wNum1)) = "" And CisFun.Val2(iB12_Hako(wNum1)) = 0 And CisFun.Val2(iB12_Syuyo(wNum1)) = 0 And _
           CisFun.Val2(iB12_Su(wNum1)) = 0 Then
        If RTrim(iB12_Bunno(wNum1)) = "" And RTrim(iB12_Noubn(wNum1)) = "" And RTrim(iB12_Hinbn(wNum1)) = "" And _
           RTrim(iB12_Seban(wNum1)) = "" And CisFun.Val2(iB12_Hako(wNum1)) = 0 And CisFun.Val2(iB12_Syuyo(wNum1)) = 0 And _
           CisFun.Val2(iB12_Su(wNum1)) = 0 Then                         ' 2009.10.30 upd : 受付№を削除

            wlc_Flg = False
            For wNum2 = wNum1 To 4
''''            If RTrim(iB12_UkeNo(wNum2)) <> "" Or RTrim(iB12_Bunno(wNum2)) <> "" Or RTrim(iB12_Noubn(wNum2)) <> "" Or RTrim(iB12_Hinbn(wNum2)) <> "" Or _
                   RTrim(iB12_Seban(wNum2)) <> "" Or CisFun.Val2(iB12_Hako(wNum2)) <> 0 Or CisFun.Val2(iB12_Syuyo(wNum2)) Or _
                   CisFun.Val2(iB12_Su(wNum2)) <> 0 Then
                If RTrim(iB12_Bunno(wNum2)) <> "" Or RTrim(iB12_Noubn(wNum2)) <> "" Or RTrim(iB12_Hinbn(wNum2)) <> "" Or _
                   RTrim(iB12_Seban(wNum2)) <> "" Or CisFun.Val2(iB12_Hako(wNum2)) <> 0 Or CisFun.Val2(iB12_Syuyo(wNum2)) Or _
                   CisFun.Val2(iB12_Su(wNum2)) <> 0 Then                ' 2009.10.30 upd : 受付№を削除

''''                    iB12_UkeNo(wNum1) = RTrim(iB12_UkeNo(wNum2))    ' 2009.10.30 del
                    iB12_Bunno(wNum1) = RTrim(iB12_Bunno(wNum2))
                    iB12_Noubn(wNum1) = RTrim(iB12_Noubn(wNum2))
                    iB12_Hinbn(wNum1) = RTrim(iB12_Hinbn(wNum2))
                    iB12_Seban(wNum1) = RTrim(iB12_Seban(wNum2))
                    iB12_Hako(wNum1) = iB12_Hako(wNum2)
                    iB12_Syuyo(wNum1) = iB12_Syuyo(wNum2)
                    iB12_Su(wNum1) = iB12_Su(wNum2)

                    iB12_UkeNo(wNum2) = ""
                    iB12_Bunno(wNum2) = ""
                    iB12_Noubn(wNum2) = ""
                    iB12_Hinbn(wNum2) = ""
                    iB12_Seban(wNum2) = ""
                    iB12_Hako(wNum2) = ""
                    iB12_Syuyo(wNum2) = ""
''''                    iB12_Su(wNum2) = ""                             ' 2009.10.30 del

                    wlc_Flg = True
                    Exit For
                End If
            Next wNum2
            If Not wlc_Flg Then Exit For
        End If
    Next wNum1
End Sub
'*----------------------------------------------------------------*
'*          伝票内容チェック
'*
'*      画面表示項目をチケット読取データの様に、文字列にして
'*                          ストアドにて内容のチェックを行う。
'*      チェック後に、ストアドからの Outputﾊﾟﾗﾒｰﾀを使用し、
'*                                              再表示をする。
'*----------------------------------------------------------------*
Private Function ContentsChk() As Boolean
    ContentsChk = False
 
    Static wcce_Obj     As Object
    
    sFlg = 0
'---------------------------------------------------------------------------' 2009.11.05 add start
    If SyoriKb = "ADD" Then
       iB1_Ku = " "
       iB1_SRTorcd = "04370"
       iB1_Ymd = Format(Now, "YYYYMMDD")
       iB1_Bin = "  "
       iB1_Gyo = " "
    End If
'
    If Trim(iB1_Hacyu) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_Title = ""
            .MB_Button = Error
            .MB_MSG(2) = "   発注者を入力して下さい              "
            .MBOX
        End With
        iB1_Hacyu.SetFocus
        Exit Function
    End If
'---------------------------------------------------------------------------' 2009.11.05 add end
'
    Static HasuChk      As Integer
'
    HasuChk = 9
    For wNum1 = 0 To 4
'
'        If CisFun.Val2(iB12_Syuyo(wNum1)) <> 0 Then
''               * 端数･箱数チェック*
'            wNumL = CisFun.Val2(iB12_Su(wNum1)) \ CisFun.Val2(iB12_Syuyo(wNum1))
'            If (CisFun.Val2(iB12_Su(wNum1)) Mod CisFun.Val2(iB12_Syuyo(wNum1))) Then
'                wNumL = wNumL + 1
'            End If
''
'            If CisFun.Val2(iB12_Hako(wNum1)) <> wNumL Then
'                With CisFun
'                    .MB_Lines = 5
'                    .MB_Title = ""
'                    .MB_Button = Error
'                    .MB_MSG(1) = "【  " & StrConv(wNum1 + 1, vbWide) & " 行目  】"
'                    .MB_MSG(3) = "      箱数エラー          "
'                    .MBOX
'                End With
'                iB12_Hako(wNum1).SetFocus
'                Exit Function
'            End If
'        End If
'
'           * 端数確認メッセージ *
        wNumL = CisFun.Val2(iB12_Syuyo(wNum1)) * CisFun.Val2(iB12_Hako(wNum1))
        If CisFun.Val2(iB12_Su(wNum1)) <> wNumL Then
            If HasuChk = 9 Then HasuChk = wNum1
        End If
'
    Next wNum1
'
    If SyoriKb = "ADD" Then
        If HasuChk <> 9 Then
            With CisFun
                .MB_Lines = 5
                .MB_Title = ""
                .MB_Button = OK_CAN
                .MB_MSG(1) = "【収容数 × 箱数 ＝ 納入数】"
                .MB_MSG(2) = "          になっていない行が有ります   "
                .MB_MSG(4) = "         処理を続行しますか？"
                If Not .MBOX Then
                      ProcHB = "B1"
                      Call SyoriName(SyoriKb)
                      Call DispChange(ProcHB)
                      iB12_Hako(HasuChk).SetFocus
                      Exit Function
                Else
                      sFlg = 1
                End If
            End With
        End If
    End If

    wGyo = 0
    For wNum1 = 0 To 4
        If Trim(iB12_Seban(wNum1)) <> "" Or Trim(iB12_Hinbn(wNum1)) <> "" Then
           wGyo = wGyo + 1
        End If
    Next wNum1
    iB14_Gyo = wGyo
    iB1_Gyo = wGyo
' ｽﾄｱﾄﾞ引渡し用　ﾃﾞｰﾀ作成
    Call CrtStoado("H")

    If RTrim(wcc_Head) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_Title = ""
            .MB_Button = Error
            .MB_MSG(2) = "  ヘッド部(発注者～機番)を入力して下さい      "   ' 2009.11.05 upd
            .MBOX
        End With
        iB1_Hacyu.SetFocus      ' 2009.11.05 upd
        Exit Function
    End If

' ｽﾄｱﾄﾞ引渡し用　ﾃﾞｰﾀ作成
    Call CrtStoado("B")

    If RTrim(wcc_Body) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_Title = ""
            .MB_Button = Error
            .MB_MSG(2) = "     明細行を入力して下さい          "
            .MBOX
        End With
'        iB12_UkeNo(0).SetFocus     ' 2009.11.02 del
        iB12_Noubn(0).SetFocus      ' 2009.11.02 add
        Exit Function
    End If
'
    If SyoriKb = "RED" Then
       gInt = 0                 ' memo : 0/QR読込
    Else
       gInt = 1                 ' memo : 1/手入力
    End If
    If ChekData(gInt, 0, 1) Then
       DoEvents
       ContentsChk = True
       Exit Function
    Else
       On Error Resume Next
       mFocusObj.SetFocus
       Set mFocusObj = Nothing
       On Error GoTo 0
    End If

End Function
'+----------------------------------------------------------------------+   ' 2009.10.30 upd : 全体的に今回使用に合わせた
'+      ストアド引渡し用　データ作成
'+----------------------------------------------------------------------+
Private Sub CrtStoado(HB As String)
'   * ヘッド部作成 *                ' 位置やデータ内容の把握のため、QRコードデータフォーマットのヘッダに合わせることとする
    If HB = "H" Then
        wcc_Head = ""
        ' 03：フォーマットID
        wcc_Head = wcc_Head & "H01"
        ' 01：行№
        If RTrim(iB1_Gyo) = "" Then
           wcc_Head = wcc_Head & Space(1)
        Else
           wcc_Head = wcc_Head & Format(RTrim(iB1_Gyo), "@")
        End If
        ' 01：発注区分
        If RTrim(iB1_Ku) = "" Then
           wcc_Head = wcc_Head & Space(1)
        Else
           wcc_Head = wcc_Head & Format(RTrim(iB1_Ku), "@")
        End If
        ' 01：自支給区分
        wcc_Head = wcc_Head & "0"
        ' 06：仕入先
        If RTrim(iB1_SRTorcd) = "" Then
           wcc_Head = wcc_Head & Space(6)
        Else
           wcc_Head = wcc_Head & Format(RTrim(iB1_SRTorcd), "!@@@@@@")
        End If
        ' 01：あき（仕入先予備）
        wcc_Head = wcc_Head & " "
        ' 05：納入先
        If RTrim(iB1_Hacyu) = "" Then
           wcc_Head = wcc_Head & Space(5)
        Else
           wcc_Head = wcc_Head & Format(RTrim(iB1_Hacyu), "!@@@@@")
        End If
        ' 02：受入
        If RTrim(iB1_Uke) = "" Then
           gSL_Select = "SELECT * FROM 取引先受入マスタ WHERE 取引先CD = '" & Trim(iB1_Hacyu) & "' AND 取引先区分=1"
           If TUMRead(gSL_Select) Then
              iB1_Uke = TUM.受入
           Else
              gSL_Select = "SELECT * FROM 受注取引先CD変換マスタ WHERE 変換元CD = '" & Trim(iB1_Hacyu) & "'"
              If JTHRead(gSL_Select) Then
                 iB1_Uke = JTH.変換先受入
              End If
           End If
        End If
        If RTrim(iB1_Uke) = "" Then
           'wcc_Head = wcc_Head & Space(2)
           wcc_Head = wcc_Head & "01"
        Else
           wcc_Head = wcc_Head & Format(RTrim(iB1_Uke), "!@@")
        End If
        ' 02：あき（納入先予備）
        wcc_Head = wcc_Head & "  "
        ' 02：機番
        If RTrim(iB1_Kibn) = "" Then
           wcc_Head = wcc_Head & Space(2)
        Else
           wcc_Head = wcc_Head & Format(RTrim(iB1_Kibn), "!@@")
        End If
        ' 03：明細行№
        wcc_Head = wcc_Head & "   "
        ' 02：あき
        wcc_Head = wcc_Head & "  "
        ' --：納入年月日（年が4桁なければ4桁にする）
        If Len(Trim(iH1_Ymd)) < 8 Then iH1_Ymd = "  " & Trim(iH1_Ymd)
        ' 02：納入年
        If RTrim(iH1_Ymd) = "" Then
           wcc_Head = wcc_Head & Space(2)
        Else
           wcc_Head = wcc_Head & Format(Mid(RTrim(iH1_Ymd), 3, 2), "!@@")
        End If
        ' 02：納入月
        If RTrim(iH1_Ymd) = "" Then
           wcc_Head = wcc_Head & Space(2)
        Else
           wcc_Head = wcc_Head & Format(Mid(RTrim(iH1_Ymd), 5, 2), "!@@")
        End If
        ' 02：納入日
        If RTrim(iH1_Ymd) = "" Then
           wcc_Head = wcc_Head & Space(2)
        Else
           wcc_Head = wcc_Head & Format(Mid(RTrim(iH1_Ymd), 7, 2), "!@@")
        End If
        ' 02：納入便
        If RTrim(iH1_Bin) = "" Then
           wcc_Head = wcc_Head & Space(2)
        Else
           wcc_Head = wcc_Head & Format(RTrim(iH1_Bin), "!@@")
        End If
        ' 02：発注年月日
        If RTrim(iB1_Ymd) = "" Then
           wcc_Head = wcc_Head & Space(6)
        Else
           wcc_Head = wcc_Head & Format(RTrim(Mid(iB1_Ymd, 3, 6)), "!@@@@@@")
        End If
        ' 02：発注便
        If RTrim(iB1_Bin) = "" Then
           wcc_Head = wcc_Head & Space(2)
        Else
           wcc_Head = wcc_Head & Format(RTrim(iB1_Bin), "!@@")
        End If
        ' 03：頁№
        wcc_Head = wcc_Head & "   "
        ' 03：総頁№
        wcc_Head = wcc_Head & "   "
    End If
    
    If HB = "B" Then
'   * ボディ部作成 *
        wcc_Body = ""
        For wNum1 = 0 To 4
'-----------------------------------------------------------------------' 2009.10.30 del start
'    '       * 受付№*
'            If RTrim(iB12_UkeNo(wNum1)) = "" Then
'                wcc_Body = wcc_Body & Space(5)
'            Else
'                wcc_Body = wcc_Body & Format(RTrim(iB12_UkeNo(wNum1)), "!@@@@@")
'            End If
'-----------------------------------------------------------------------' 2009.10.30 del end
    '       * 納入番号*
            If RTrim(iB12_Noubn(wNum1)) = "" Then
                wcc_Body = wcc_Body & Space(5)
            Else
                wcc_Body = wcc_Body & Format(RTrim(iB12_Noubn(wNum1)), "!@@@@@")
            End If
    '       * 分納*
            If RTrim(iB12_Bunno(wNum1)) = "" Then
                wcc_Body = wcc_Body & Space(1)
            Else
                wcc_Body = wcc_Body & Format(RTrim(iB12_Bunno(wNum1)), "@")
            End If
    '       * 背番号*
            If RTrim(iB12_Seban(wNum1)) = "" Then
                wcc_Body = wcc_Body & Space(4)
            Else
                wcc_Body = wcc_Body & Format(RTrim(iB12_Seban(wNum1)), "!@@@@")
            End If
    '       * 品番 *
            If RTrim(iB12_Hinbn(wNum1)) = "" Then
                wcc_Body = wcc_Body & Space(13)                                             ' 2009.10.30 upd : 桁数 14-->13
            Else
                wcc_Body = wcc_Body & Format(RTrim(iB12_Hinbn(wNum1)), "!@@@@@@@@@@@@@")    ' 2009.10.30 upd : 桁数 14-->13
            End If
            ' 02：あき
            wcc_Body = wcc_Body & "         "
    '       * 収容数 *
            If RTrim(iB12_Syuyo(wNum1)) = "" Then
               wcc_Body = wcc_Body & Space(5)
            Else
               wcc_Body = wcc_Body & Format(CisFun.Val2(iB12_Syuyo(wNum1)), "@@@@@")
            End If
    '       * 箱数 *
            If RTrim(iB12_Hako(wNum1)) = "" Then
               wcc_Body = wcc_Body & Space(3)
            Else
               wcc_Body = wcc_Body & Format(CisFun.Val2(iB12_Hako(wNum1)), "@@@")
            End If
    '       * 納入数 *
            If RTrim(iB12_Su(wNum1)) = "" Then
               wcc_Body = wcc_Body & Space(6)
            Else
               wcc_Body = wcc_Body & Format(CisFun.Val2(iB12_Su(wNum1)), "@@@@@@")
            End If
    '       * 打切フラグ *
            If RTrim(iB12_KiriF(wNum1)) = "" Then
               wcc_Body = wcc_Body & Space(1)
            Else
               wcc_Body = wcc_Body & Format(CisFun.Val2(iB12_KiriF(wNum1)), "@")
            End If
    '       * 打切残数 *
            If RTrim(iB12_KiriZ(wNum1)) = "" Then
               wcc_Body = wcc_Body & Space(6)
            Else
               wcc_Body = wcc_Body & Format(CisFun.Val2(iB12_KiriZ(wNum1)), "@@@@@@")
            End If
        Next wNum1
    End If
End Sub
'************************************************************
'*
'*          ＱＲ読取内容画面表示
'*
'************************************************************
Private Sub ReadDsp(RData As String)

    Dim wS      As String
    Dim wN      As Integer

'    If RTrim(iH1_Ymd) = "" Then
'        iH1_Ymd = Mid(RData, 40, 8)
'    End If
    wS = Mid(RData, 30, 2)
    If CisFun.Val2(wS) > 90 Then
        wS = "19"
    Else
        wS = "20"
    End If
    iH1_Ymd = wS & Mid(RData, 30, 6)
'    If RTrim(iH1_Bin) = "" Then
       iH1_Bin = Mid(RData, 36, 2)          'Mid(RData, 96, 2)
'    End If

'区
    iB1_Ku = Mid(RData, 5, 1)
'♯
'    iB1_Sharp = ""                          ' Mid(RData, 2, 1)                         ' 2009.10.30 del
'仕入先
'    iB1_SRTorcd = Trim(Mid(RData, 7, 6))    ' Mid(RData, 48, 4) & Mid(RData, 60, 1)    ' 2009.10.30 del
    iB1_SRTorcd = Replace(Trim(Mid(RData, 7, 6)), "-", "")                              ' 2009.10.30 add
'発注者（納入先）
    iB1_Hacyu = Mid(RData, 14, 5)           ' Mid(RData, 16, 4) & Mid(RData, 28, 1)
    If RTrim(iB1_Hacyu) = "" Then
       iB1_Hacyu = RTrim(wAkitaCode)
    End If
'受入
    iB1_Uke = Mid(RData, 19, 2)             ' Mid(RData, 32, 2)

'---------------------------------------------------------------------------------------' 2009.10.30 add start <<作成中>>
''''    If Trim(iB1_Uke) = "" Then
''''       gSL_Select = ""
''''       gSL_Select = gSL_Select & "SELECT MIN(受入) FROM 品番取引先マスタ"
''''       gSL_Select = gSL_Select & " WHERE 品番  '" & iH1_Hinbn & "' and 取引先='" & iH1_Torcd & "' --and 品目=0 "
''''    End If
'---------------------------------------------------------------------------------------' 2009.10.30 add end   <<作成中>>

'機番
    iB1_Kibn = Mid(RData, 23, 2)            ' Mid(RData, 36, 3)
'月日便
''    iB1_YY = Mid(RData, 38, 2)              ' 2009.11.05 add      ' 2009.11.05 del
''    iB1_MM = Mid(RData, 40, 2)              ' Mid(RData, 44, 2)   ' 2009.11.05 del
''    iB1_DD = Mid(RData, 42, 2)              ' Mid(RData, 46, 2)   ' 2009.11.05 del
    iB1_Ymd = "20" & Mid(RData, 38, 6)        ' 2009.11.05 add
    iB1_Bin = Mid(RData, 44, 2)
'行
    iB1_Gyo = Mid(RData, 4, 1)              ' ""
'**BODY**
    wN = 54
    For wI = 0 To 4
        If RTrim(Mid(RData, 63 + (wI * wN), 13)) <> "" Then     ' 品番が空白でなければ処理する
    '受付NO
'           iB12_UkeNo(wI) = Space(5)                            ' Mid(RData, 130 + (wI * 74), 5)   ' 2009.10.30 del
    '納入番号
           ' 分納回数
           iB12_Bunno(wI) = Mid(RData, 58 + (wI * wN), 1)       ' Mid(RData, 82 + (wI * 74), 1)
           ' 納入番号（発注納番）
           iB12_Noubn(wI) = Mid(RData, 52 + (wI * wN), 5)       ' Mid(RData, 83 + (wI * 74), 5)
    '背番号
           iB12_Seban(wI) = Mid(RData, 59 + (wI * wN), 4)       ' Mid(RData, 127 + (wI * 74), 3)
    '品番
           iB12_Hinbn(wI) = Mid(RData, 63 + (wI * wN), 13)      ' Mid(RData, 68 + (wI * 74), 14)
    '収容数
           iB12_Syuyo(wI) = Mid(RData, 85 + (wI * wN), 5)       ' Mid(RData, 112 + (wI * 74), 5)
    '箱数
           iB12_Hako(wI) = Mid(RData, 90 + (wI * wN), 3)        ' Mid(RData, 117 + (wI * 74), 3)
    '納入数
           iB12_Su(wI) = Mid(RData, 93 + (wI * wN), 6)          ' Mid(RData, 90 + (wI * 74), 6)
    '打切フラグ
           iB12_KiriF(wI) = Mid(RData, 99 + (wI * wN), 1)
    '打切残数
           iB12_KiriZ(wI) = Mid(RData, 100 + (wI * wN), 6)
        End If
    Next wI
    
    
End Sub
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+    追加 処理          +
'+----------------------+
Private Function DBPut()
    DBPut = False
'  ( 確認メッセージ )
    If Not DBCheckMsg Then Exit Function
 
    On Error GoTo DBPut_Err

' ｽﾄｱﾄﾞ引渡し用　データ作成
    Call CrtStoado("H")
    Call CrtStoado("B")

    If SyoriKb = "RED" Then
       gInt = 0
    Else
       gInt = 1
    End If
    If Not ChekData(0, 1, 1) Then
        DoEvents
    End If
    
    If SyoriKb = "RED" Then
       H1lb_OCRCnt = Format(CisFun.Val2(H1lb_OCRCnt) + 1, "#,###")
    Else
       iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "#,###")
    End If
    
    On Error GoTo 0
    
    DBPut = True
    Exit Function
DBPut_Err:
    Call CisFun.ErrorBox("更新エラー")
    End
End Function
'+------------------------------------+
'+    処理確認メッセージ（DBUpDate)
'+------------------------------------+
Private Function DBCheckMsg()
    DBCheckMsg = False
    
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
         .MB_Lines = 3
         .MB_Title = SyoriNM & "処理"
         .MB_Button = OK_CAN
         .MB_MSG(1) = "    【 " & SyoriNM & "処理確認 】      "
         
         If Not .MBOX Then Exit Function
    End With
    
    DBCheckMsg = True
End Function
'==========================================================='
'   終了処理（入力データがあれば後方処理を行う）
'==========================================================='
Private Sub EndProc()

'   << 入力件数確認 >>
    gSL_Select = ""
    gSL_Select = gSL_Select & "select count(*) 件数"
    gSL_Select = gSL_Select & "  from 受注入力ワーク"
    gSL_Select = gSL_Select & " where 端末番号 = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and"
    gSL_Select = gSL_Select & " ( 出荷指示書NO = 0"
    gSL_Select = gSL_Select & "    or 累積区分 = 0"
    gSL_Select = gSL_Select & " )"
    If JNWRead(gSL_Select, 1) Then
        If JNW.件数 > 0 Then
        '   出荷指示NO採番
'*--------------------------------------------------<< 2004/06/24 Insert Start >>
            gCHT_DelKbn53 = 0
'*--------------------------------------------------<< 2004/06/24 Insert End >>
            CXJ0110.Show vbModal
            Unload CXJ0110
            Set CXJ0110 = Nothing
            DoEvents
        '   出荷指示書発行
            Call CXJ0120LPrint(0)
        '   自社伝票発行
            Call CXJ0140LPrint(0)
        '   自社伝票ﾌｧｲﾙ出力
            Call CXJ0150TFile(0)
        '   ノーかんばんエフ発行
            Call CXJ0130LPrint(0)
        End If
    End If

End Sub
'+---------------------------------------+
'+      名称マスタ初期設定
'+---------------------------------------+
Private Sub MeisyoInit()
'   <><><><><> 受注区分 <><><><><>
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "受注区分"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "通常"
        .作成者 = gTanto
        .制御区分 = "X"
    End With
    If Not SYMInitCreate(True) Then
        With SYMCreate
            .値 = "2":  .値名称 = "特配":       Call SYMInitCreate
            .値 = "3":  .値名称 = "臨時":       Call SYMInitCreate
            .値 = "4":  .値名称 = "スポット":   Call SYMInitCreate
            .値 = "5":  .値名称 = "指示":       Call SYMInitCreate
            .値 = "6":  .値名称 = "号試":       Call SYMInitCreate
            .値 = "7":  .値名称 = "代替":       Call SYMInitCreate
            .値 = "8":  .値名称 = "その他":     Call SYMInitCreate
        End With
    End If
'   >> 受注区分の表示
    gSL_Select = ""
    gSL_Select = gSL_Select & "select 値"
    gSL_Select = gSL_Select & "     , 値名称"
    gSL_Select = gSL_Select & "  from 名称マスタ"
    gSL_Select = gSL_Select & " where 区分名称 = '受注区分'"
    gSL_Select = gSL_Select & "   and 区分タイプ = ''"
    gSL_Select = gSL_Select & " order by 値"
    If SYMRead(gSL_Select, 0) Then
        Do Until Not SYM_RDSTS
            Select Case SYM.値
                Case "1":   H1Op_JKbn(1).Caption = SYM.値名称:  H1Op_JKbn(1).Tag = SYM.値
                Case "5":   H1Op_JKbn(2).Caption = SYM.値名称:  H1Op_JKbn(2).Tag = SYM.値
                Case "6":   H1Op_JKbn(3).Caption = SYM.値名称:  H1Op_JKbn(3).Tag = SYM.値
                Case "4":   H1Op_JKbn(4).Caption = SYM.値名称:  H1Op_JKbn(4).Tag = SYM.値
                Case "3":   H1Op_JKbn(5).Caption = SYM.値名称:  H1Op_JKbn(5).Tag = SYM.値
                Case "2":   H1Op_JKbn(6).Caption = SYM.値名称:  H1Op_JKbn(6).Tag = SYM.値
                Case "7":   H1Op_JKbn(7).Caption = SYM.値名称:  H1Op_JKbn(7).Tag = SYM.値
                Case "8":   H1Op_JKbn(8).Caption = SYM.値名称:  H1Op_JKbn(8).Tag = SYM.値
            End Select
            Call SYMReadNext
        Loop
    End If
    Call SYMClose
End Sub
'+--------------------------------------------+
'+      読  取  処  理(ﾃｽﾄ用ﾃｷｽﾄﾌｧｲﾙREAD)      +
'+--------------------------------------------+
Private Sub Test_OnComm()
    Static c_FNo    As Byte
    Static c_STS    As Integer
    Static c_Data   As String

    c_FNo = FreeFile
    Open gIniExe & "TBTiketo.txt" For Input Access Read As #c_FNo
    
    Do Until EOF(c_FNo)
        Line Input #c_FNo, c_Data
    
        Call ReadDsp(c_Data)
             
        With CisFun
             .MB_Lines = 3
             .MB_Title = "【Data Check】"
             .MB_Button = OK_CAN
             .MB_MSG(1) = "Data Write. Are You OK??"
             
             If Not .MBOX Then Exit Sub
        End With
        
        If ContentsChk Then
'           If Not DBPut Then Exit Sub
           H1lb_OCRCnt = Format(CisFun.Val2(H1lb_OCRCnt) + 1, "###0")
           Call HeadBodyClear("B")
           iH1_Ymd = ""
           iH1_Bin = ""
           DoEvents
        Else
           Exit Sub
        End If
    Loop
    
    Close #c_FNo
    
End Sub
'******************************************************************
'*      初期処理
'******************************************************************
Private Function InitProc() As Boolean
    InitProc = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select *"
    gSL_Select = gSL_Select & "  from 受注入力ワーク"
    gSL_Select = gSL_Select & " where 端末番号 = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and 累積区分 = 0"
    If JNWRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 6
            .MB_MSG(1) = "    当端末処理分の受注未累積データが存在します。            "
            .MB_MSG(4) = "      『受注画面入力（号口）』処理を終了しますか？"
            .MB_MSG(5) = "      いいえの場合は、未累積データを削除し続行します"
            .MB_Title = "未累積データの確認"
            .MB_Button = Yes_No
            If .MBOX Then Exit Function
        End With
    End If

    With CisDB
        On Error GoTo InitProc_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
'       << 受注入力テーブル >>
        .SQL = ""
        .SQL = .SQL & "delete"
        .SQL = .SQL & "  from 受注入力ワーク"
        .SQL = .SQL & " where 端末番号 = substring(host_name(),1,20)"
        .DBExec
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        On Error GoTo 0
    End With

    InitProc = True
    Exit Function
InitProc_Err:
    CisFun.ErrorBox
    End
End Function

''+-------------------------------------+
''+      Debug用 SQL記録                +
''+-------------------------------------+
Private Sub Test_LogOutput(str As String, Optional Msg As Integer = 0)
'Exit Sub   '=========================================== 実行しないときはこの行を生かすこと =========================
    Dim dt01 As String
    Dim dt02 As String
    Dim f_No As Integer
    
    If Msg <> 0 Then MsgBox (str)
    dt01 = CStr(DateTime.Now)
    dt02 = Mid(dt01, 1, 4) & "." & _
           Mid(dt01, 6, 2) & "." & _
           Mid(dt01, 9, 2)

    f_No = FreeFile                             'ファイル番号の取得
    Open "D:\TestLog\SQL_" & dt02 & ".log" _
          For Append As #f_No                   'ファイルを追加モードで開く
    Print #f_No, "[" & Me.Name & "] " & dt01 & _
                 " ----------------------------------------" & Chr(13) & Chr(10) & _
                 str & Chr(13) & Chr(10)        'ファイルへ書き込む"
    Close #f_No                                 'ファイルを閉じる
End Sub
