VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{1BA4D453-0150-11CE-89B0-0000C037528B}#3.2#0"; "PDQCom32.OCX"
Begin VB.Form CMJ0020 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "前工程JAMA 入荷実績入力 [個別]"
   ClientHeight    =   10290
   ClientLeft      =   4080
   ClientTop       =   3495
   ClientWidth     =   14580
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
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
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   10290
   ScaleWidth      =   14580
   WindowState     =   2  '最大化
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
      Height          =   3960
      Left            =   2850
      ScaleHeight     =   3900
      ScaleWidth      =   7230
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1815
      Width           =   7290
      Begin Cis3D_v60.CIS3D CIS3D23 
         Height          =   690
         Left            =   1905
         Top             =   45
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   1217
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
      End
      Begin Cis3D_v60.CIS3D CIS3D22 
         Height          =   60
         Left            =   60
         Top             =   735
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   106
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
      End
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   450
         Left            =   60
         Top             =   795
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   794
         ForeColor       =   0
         Caption         =   "納入先"
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin CisText_V60.CisText iB1_Nonyu 
            Height          =   330
            Left            =   1605
            TabIndex        =   7
            Top             =   75
            Width           =   690
            _ExtentX        =   1217
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
            Text            =   "XXXX"
            MaxLength       =   4
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   450
         Left            =   60
         Top             =   1245
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   794
         ForeColor       =   0
         Caption         =   "納入先事業所"
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   195
            Left            =   1890
            Top             =   150
            Width           =   120
            _ExtentX        =   212
            _ExtentY        =   344
            Caption         =   "-"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
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
         Begin CisText_V60.CisText iB1_NonyuKj 
            Height          =   330
            Left            =   1605
            TabIndex        =   8
            Top             =   75
            Width           =   255
            _ExtentX        =   450
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
            Text            =   "X"
            MaxLength       =   1
         End
         Begin CisText_V60.CisText iB1_Ukeir 
            Height          =   330
            Left            =   2025
            TabIndex        =   9
            Top             =   75
            Width           =   360
            _ExtentX        =   635
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
            Text            =   "XX"
            MaxLength       =   2
         End
         Begin Cis3D_v60.CIS3D B1lb_JUkeir 
            Height          =   285
            Left            =   3075
            Top             =   90
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            BackColor       =   12648384
            Caption         =   "XX"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   450
         Left            =   60
         Top             =   1695
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   794
         ForeColor       =   0
         Caption         =   "納入指示便"
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin CisText_V60.CisText iB1_NBin 
            Height          =   330
            Left            =   1605
            TabIndex        =   10
            Top             =   75
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   582
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
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   450
         Left            =   60
         Top             =   2145
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   794
         ForeColor       =   0
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin CisText_V60.CisText iB1_Kiban 
            Height          =   330
            Left            =   1605
            TabIndex        =   11
            Top             =   75
            Width           =   345
            _ExtentX        =   609
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
            Text            =   "XX"
            MaxLength       =   2
            cDataReplace    =   1
            cFaZero         =   0
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D15 
         Height          =   450
         Index           =   0
         Left            =   60
         Top             =   2595
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   794
         ForeColor       =   0
         Caption         =   "カード＃"
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin CisText_V60.CisText iB1_Card 
            Height          =   330
            Left            =   1605
            TabIndex        =   12
            Top             =   75
            Width           =   240
            _ExtentX        =   423
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
            Text            =   "X"
            MaxLength       =   1
            cDataReplace    =   1
            cFaZero         =   0
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D16 
         Height          =   810
         Left            =   60
         Top             =   3045
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   1429
         ForeColor       =   0
         Caption         =   "支給元"
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D B1lb_KTornm 
            Height          =   345
            Left            =   3285
            Top             =   60
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   609
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D CIS3D17 
            Height          =   195
            Left            =   2265
            Top             =   150
            Width           =   120
            _ExtentX        =   212
            _ExtentY        =   344
            Caption         =   "-"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
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
         Begin CisText_V60.CisText iB1_SikyuK 
            Height          =   330
            Left            =   1605
            TabIndex        =   13
            Top             =   75
            Width           =   600
            _ExtentX        =   1058
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
            Text            =   "XXXX"
            MaxLength       =   4
         End
         Begin CisText_V60.CisText iB1_Sikyu1 
            Height          =   330
            Left            =   2670
            TabIndex        =   14
            Top             =   75
            Width           =   615
            _ExtentX        =   1085
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
            Text            =   "XXXX"
            MaxLength       =   4
         End
         Begin CisText_V60.CisText iB1_Sikyu2 
            Height          =   330
            Left            =   2670
            TabIndex        =   15
            Top             =   435
            Width           =   615
            _ExtentX        =   1085
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
            Text            =   "XXXX"
            MaxLength       =   4
         End
         Begin Cis3D_v60.CIS3D CIS3D21 
            Height          =   225
            Left            =   2415
            Top             =   480
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   397
            Caption         =   "②"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
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
         Begin Cis3D_v60.CIS3D CIS3D25 
            Height          =   225
            Left            =   2415
            Top             =   135
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   397
            Caption         =   "①"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
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
         Begin Cis3D_v60.CIS3D CIS3D26 
            Height          =   225
            Left            =   1095
            Top             =   135
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   397
            Caption         =   "体系"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
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
         Begin Cis3D_v60.CIS3D B1lb_JKTorcd 
            Height          =   285
            Left            =   5985
            Top             =   90
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   503
            BackColor       =   12648384
            Caption         =   "XXXXXXX"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B1lb_JDaikoK 
            Height          =   285
            Left            =   5985
            Top             =   420
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   503
            BackColor       =   12648384
            Caption         =   "XXXXXXX"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   450
         Left            =   3615
         Top             =   795
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   794
         ForeColor       =   0
         Caption         =   "受注者"
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D B1lb_JSiire 
            Height          =   285
            Left            =   2460
            Top             =   90
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   503
            BackColor       =   12648384
            Caption         =   "XXXXXXX"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin CisText_V60.CisText iB1_Siire 
            Height          =   330
            Left            =   1605
            TabIndex        =   16
            Top             =   75
            Width           =   690
            _ExtentX        =   1217
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
            Text            =   "XXXX"
            MaxLength       =   4
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   450
         Left            =   3615
         Top             =   1245
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   794
         ForeColor       =   0
         Caption         =   "受注者事業所"
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D CIS3D12 
            Height          =   195
            Left            =   1890
            Top             =   150
            Width           =   120
            _ExtentX        =   212
            _ExtentY        =   344
            Caption         =   "-"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
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
         Begin CisText_V60.CisText iB1_SiireKj 
            Height          =   330
            Left            =   1605
            TabIndex        =   17
            Top             =   75
            Width           =   255
            _ExtentX        =   450
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
            Text            =   "X"
            MaxLength       =   1
         End
         Begin CisText_V60.CisText iB1_Syukaba 
            Height          =   330
            Left            =   2025
            TabIndex        =   18
            Top             =   75
            Width           =   480
            _ExtentX        =   847
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
            Text            =   "XXX"
            MaxLength       =   3
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   690
         Left            =   5640
         Top             =   45
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   1217
         ForeColor       =   0
         Caption         =   "総ページ"
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
         Begin CisText_V60.CisText iB1_TPage 
            Height          =   330
            Left            =   645
            TabIndex        =   6
            Top             =   300
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   582
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
      Begin Cis3D_v60.CIS3D CIS3D19 
         Height          =   690
         Left            =   3615
         Top             =   45
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   1217
         ForeColor       =   0
         Caption         =   "発行日"
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
         Begin CisYMD_With_Btn.CisYMDwB iB1_HYmd 
            Height          =   315
            Left            =   150
            TabIndex        =   5
            Top             =   300
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
      End
      Begin Cis3D_v60.CIS3D CIS3D20 
         Height          =   690
         Left            =   60
         Top             =   45
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   1217
         ForeColor       =   0
         Caption         =   "発注区分"
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
         Begin Cis3D_v60.CIS3D B1lb_HKbn 
            Height          =   315
            Left            =   450
            Top             =   300
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            BackColor       =   16777152
            Caption         =   "XXXXXXXX"
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
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   60
         End
         Begin CisText_V60.CisText iB1_HKbn 
            Height          =   330
            Left            =   210
            TabIndex        =   4
            Top             =   300
            Width           =   240
            _ExtentX        =   423
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
            Text            =   "X"
            MaxLength       =   1
            cDataReplace    =   1
            cFaZero         =   0
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D24 
         Height          =   1350
         Left            =   3615
         Top             =   1695
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   2381
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin Cis3D_v60.CIS3D B1_Sosin 
            Height          =   435
            Left            =   240
            Top             =   825
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   767
            BackColor       =   16761024
            Caption         =   "送信日"
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
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   60
            Begin Cis3D_v60.CIS3D B1lb_TNSYmd 
               Height          =   360
               Left            =   885
               Top             =   45
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   635
               BackColor       =   16761024
               Caption         =   "YYYY/MM/DD"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   16761024
               cBoderColor2    =   16761024
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cBoderStyle     =   2
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D27 
            Height          =   540
            Left            =   120
            Top             =   120
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   953
            Caption         =   "※納入先･納入先事業所･受注者・受注者　事業所は支給元のｺｰﾄﾞです。"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
         End
      End
   End
   Begin Cis3D_v60.CIS3D H3_Area1 
      Height          =   1110
      Left            =   6720
      Top             =   660
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   1958
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
      Begin VB.TextBox QRDummy 
         Height          =   330
         Left            =   0
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   0
         Width           =   240
      End
      Begin Cis3D_v60.CIS3D QRMsg 
         Height          =   990
         Left            =   75
         Top             =   75
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1746
         BackColor       =   12648384
         ForeColor       =   16711680
         Caption         =   "QRリーダー読取中"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   7
         cAlingnment     =   7
      End
   End
   Begin Cis3D_v60.CIS3D ErrMsg 
      Height          =   465
      Left            =   10230
      Top             =   1905
      Visible         =   0   'False
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   820
      BackColor       =   16761087
      Caption         =   "エラー"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   14.25
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
      cPositionX      =   90
   End
   Begin PdqcommLib.PDQComm PDQComm1 
      Height          =   480
      Left            =   825
      TabIndex        =   32
      Top             =   1275
      Width           =   480
      _Version        =   196610
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.26
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
   Begin VB.PictureBox H1_Area1 
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
      Height          =   1110
      Left            =   2865
      ScaleHeight     =   1050
      ScaleWidth      =   3705
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   660
      Width           =   3765
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   960
         Left            =   2115
         Top             =   45
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   1693
         ForeColor       =   16711680
         Caption         =   "納 番"
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
         Begin CisText_V60.CisText iH1_Nouno 
            Height          =   360
            Left            =   120
            TabIndex        =   1
            Top             =   420
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   635
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
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   960
         Left            =   3105
         Top             =   45
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   1693
         ForeColor       =   16711680
         Caption         =   "頁"
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
         Begin CisText_V60.CisText iH1_Page 
            Height          =   360
            Left            =   75
            TabIndex        =   2
            Top             =   420
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   635
            cFormat         =   "##"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      Begin Cis3D_v60.CIS3D CIS3D10 
         Height          =   960
         Left            =   75
         Top             =   45
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1693
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisYMD_With_Btn.CisYMDwB iH1_NYmd 
            Height          =   360
            Left            =   90
            TabIndex        =   0
            Top             =   420
            Width           =   1860
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
            Object.Height          =   360
            Object.Width           =   1860
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
      End
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   10230
      ScaleHeight     =   1260
      ScaleWidth      =   4305
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   480
      Width           =   4365
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   2595
         Top             =   630
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         Left            =   3450
         Top             =   630
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         Left            =   3450
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         Left            =   30
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
      Begin CisBtn_60.CisBtn PB_CHG 
         Height          =   585
         Left            =   1740
         Top             =   630
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   10
         BtnCaption      =   "切替"
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   2595
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   4
         BtnCaption      =   "削除"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         Left            =   885
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
      Begin CisBtn_60.CisBtn PB_MNT 
         Height          =   585
         Left            =   1740
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   8
         BtnCaption      =   "修正"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
      Begin CisBtn_60.CisBtn PB_NEXT 
         Height          =   585
         Left            =   885
         Top             =   630
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   5
         BtnCaption      =   "前進"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         Left            =   30
         Top             =   630
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   6
         BtnCaption      =   "後退"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
   Begin VB.PictureBox B1_Area2 
      Height          =   4200
      Left            =   2820
      ScaleHeight     =   4140
      ScaleWidth      =   7260
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5790
      Width           =   7320
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   4050
         Left            =   45
         Top             =   45
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   7144
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin vsFlexLib.vsFlexArray VsGrid1 
            Height          =   3735
            Left            =   75
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   150
            Width           =   7020
            _Version        =   196608
            _ExtentX        =   12382
            _ExtentY        =   6588
            _StockProps     =   228
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "№|品　番                   |背番号|収容数 |箱 数|納入数  |"
            Rows            =   11
            BackColor       =   12648384
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   0
            ScrollBars      =   0
            RowHeightMin    =   330
            GridLines       =   3
            SelectionMode   =   1
            BackColorAlternate=   12648384
            Begin Cis3D_v60.CIS3D Input_Area 
               Height          =   375
               Left            =   330
               Top             =   330
               Width           =   6630
               _ExtentX        =   11695
               _ExtentY        =   661
               BackColor       =   0
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
               Begin CisText_V60.CisText iB1_Seban 
                  Height          =   375
                  Left            =   3090
                  TabIndex        =   21
                  Top             =   0
                  Width           =   825
                  _ExtentX        =   1455
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
                  IMEMode         =   2
                  cDataReplace    =   1
               End
               Begin CisText_V60.CisText iB1_Hinbn 
                  Height          =   375
                  Left            =   15
                  TabIndex        =   20
                  Top             =   0
                  Width           =   3135
                  _ExtentX        =   5530
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
                  Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   25
                  cDataReplace    =   1
               End
               Begin CisText_V60.CisText iB1_Maisu 
                  Height          =   360
                  Left            =   4830
                  TabIndex        =   23
                  Top             =   0
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   635
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
                  Text            =   "#####"
                  MaxLength       =   5
                  cDataType       =   1
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#####"
                  cILength        =   4
               End
               Begin CisText_V60.CisText iB1_HSuryo 
                  Height          =   360
                  Left            =   5520
                  TabIndex        =   24
                  Top             =   0
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   635
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   7
               End
               Begin CisText_V60.CisText iB1_Syuyo 
                  Height          =   360
                  Left            =   3900
                  TabIndex        =   22
                  Top             =   0
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   635
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
                  Text            =   "#####"
                  MaxLength       =   5
                  cDataType       =   1
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#####"
                  cILength        =   4
               End
            End
         End
         Begin CisText_V60.CisText B_STabDummy 
            Height          =   360
            Left            =   165
            TabIndex        =   19
            Top             =   195
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
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
         Begin CisText_V60.CisText B_TabDummy 
            Height          =   360
            Left            =   570
            TabIndex        =   25
            Top             =   195
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
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
   End
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'ｵﾌ固定
      Left            =   960
      TabIndex        =   27
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
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 前工程発行JAMA 入荷実績入力 [個別]  】"
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
      cPositionX      =   3000
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'なし
         Height          =   345
         Index           =   1
         Left            =   10575
         ScaleHeight     =   345
         ScaleWidth      =   3675
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   30
         Width           =   3675
         Begin Cis3D_v60.CIS3D CIS3D15 
            Height          =   345
            Index           =   1
            Left            =   2445
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "削除"
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
            Left            =   15
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "追加"
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
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "修正"
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
         Left            =   14280
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
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   65535
         Caption         =   "XXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cBoderWidth     =   3
         cFont3D         =   2
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D K_Sykbnm2 
         Height          =   375
         Left            =   1200
         Top             =   0
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   128
         Caption         =   "XXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cBoderWidth     =   3
         cFont3D         =   2
         cAlingnment     =   7
      End
   End
End
Attribute VB_Name = "CMJ0020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   前工程発行JAMA 入荷実績入力 [個別]
'**       フォームID    :   CMJ0020
'**       処理概要      :
'**
'**       作  成  日    :   2007/07/27  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)
    
    Dim mSL_Select              As String
    
    Dim CisVsGrid               As New CisVsGrid3
    
    Dim mHeadProc               As String
    Dim mSaveHeadProc           As String
    
    Dim ActObj          As Object
    Dim ErrObj          As Object
    Dim ActRow          As Integer
    
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()

    If Rv_NCall = "" Then
    '   #-------------------#
    '   # 二 重 起 動 防 止  #
    '   #-------------------#
        If Not Dupli_Start Then End
    '   #-------------------#
    '   # 初 期 内 容 取 得  #
    '   #-------------------#
        If Not IniGet Then End
        '   ( QR Reader 環境 )
        Set QRr_Object = PDQComm1
        Set QRr_MsgObject = QRMsg
        Set QRr_ErrMsgObj = ErrMsg
        Call QRInit
    End If
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    QRDummy.Left = -1000
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 1000   ' 処理区分点滅間隔
    Timer2.Enabled = True   ' 処理区分点滅禁止
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = VsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 10
        .InitGet = False
        .Init
    End With
    Call GridInit
    
    Call MeisyoInit
'   #-------------------#
'   #  項 目 初 期 化   #
'   #-------------------#
    Call CisFrm.InitFld
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
    
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    K_Sykbnm2.Tag = "0"
    Call PB_CHG_Click
    
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = "ADD"
    Call SyoriName(SyoriKB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    If Rv_NCall = "" Then
       Me.Show
    Else
    End If
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    FormAct = Not FormAct


End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Rv_NCall = "" Then
       Cancel = 1
    End If
    Call PB_END_Click
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》                 +
'+-------------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+-------------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》         +
'+-------------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Key_Acc = False
    
    Select Case Key_Code
           Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click    '【新規】
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    '【修正】
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF6:       If PB_CHG.Visible Then Call PB_CHG_Click        '【切替】
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_Can_Click    '【取消】
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyPageUp:    If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
           Case vbKeyPageDown:  If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                      '【入力】
                                    Call PB_ENT_Click
                                End If
           Case Else:           Exit Function
    End Select
    
    Key_Acc = True
End Function
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    If Rv_NCall = "" Then
        Call CisDB.DBDISConnect                            ' ﾃﾞｰﾀﾍﾞｰｽ切断
        Call CisFrm.UnLoadDisp                          ' ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
        End
    Else
        Timer1.Enabled = False
        Me.Hide
    End If
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
Private Sub PB_Can_Click()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "B1" Then
        
        ProcHB = mHeadProc
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        
        GoTo PB_CAN_Ed
    End If
    
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    新  規  キ ー(F1)       +
'+----------------------------+
Private Sub PB_ADD_Click()
    If SyoriKB <> "ADD" Then
        SyoriKB = "ADD"
        Call SyoriName(SyoriKB)
    End If
    If ProcHB = "H1" Then iH1_NYmd.SetFocus
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If SyoriKB <> "MNT" Then
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
    End If
    If ProcHB = "H1" Then iH1_NYmd.SetFocus
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
        SyoriKB = "DEL"
    Else
        SyoriKB = ""
    End If
    Call SyoriName(SyoriKB)
    If ProcHB = "H1" Then iH1_NYmd.SetFocus
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    
    gSL_Select = "SELECT * FROM 受給検収テーブル"
    gSL_Select = gSL_Select & " WHERE  納入日 + 納入番号 + Cast(納品書頁 as Char(2)) < "
    gSL_Select = gSL_Select & " '" & RTrim$(iH1_NYmd) & RTrim$(iH1_Nouno) & Format(iH1_Page, "0") & "'"
    gSL_Select = gSL_Select & " ORDER BY 納入日 + 納入番号 + Cast(納品書頁 as Char(2)) DESC "
    
    If Not JYKRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    先頭レコードです            "
            .MB_Title = "【前頁】"
            .MB_Button = OK
            .MBOX
        End With
        GoTo PB_Back_ED
    End If
    
    iH1_NYmd = JYK.納入日
    iH1_Nouno = JYK.納入番号
    iH1_Page = JYK.納品書頁
    
    Call DBInput("INV")
PB_Back_ED:
    
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    
    gSL_Select = "SELECT * FROM 受給検収テーブル"
    gSL_Select = gSL_Select & " WHERE  納入日 + 納入番号 + Cast(納品書頁 as Char(2)) > "
    gSL_Select = gSL_Select & " '" & RTrim$(iH1_NYmd) & RTrim$(iH1_Nouno) & Format(iH1_Page, "0") & "'"
    gSL_Select = gSL_Select & " ORDER BY 納入日 + 納入番号 + Cast(納品書頁 as Char(2)) ASC "
    
    If Not JYKRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    最終レコードです            "
            .MB_Title = "【次頁】"
            .MB_Button = OK
            .MBOX
        End With
        GoTo PB_Next_ED
    End If
    
    iH1_NYmd = JYK.納入日
    iH1_Nouno = JYK.納入番号
    iH1_Page = JYK.納品書頁
    
    Call DBInput("INV")
PB_Next_ED:
    Call CisFrm.MousePT(1)
End Sub
'+------------------+
'+    切替   (F6)   +
'+------------------+
Private Sub PB_CHG_Click()
    gLong = CisFun.Val2(K_Sykbnm2.Tag)
    
    gLong = gLong + 1
    If gLong > 2 Then gLong = 1
    
    mSaveHeadProc = "H1"
    If IRN_NKRead = 1 Then
       mSaveHeadProc = mSaveHeadProc & "H3"
    Else
       mSaveHeadProc = mSaveHeadProc & "  "
    End If
    
    mHeadProc = ""
    gInt = 1
    Do Until Trim(mHeadProc) <> ""
        mHeadProc = Mid(mSaveHeadProc, gLong + (gLong - 1), 2)
        If Trim(mHeadProc) <> "" Then
           K_Sykbnm2.Tag = gLong
           Exit Do
        End If
        gLong = gLong + 1
        If gLong > 2 Then gLong = 1
    Loop
    
    H3_Area1.Visible = False
    Select Case mHeadProc
        Case "H3"
            K_Sykbnm2.Caption = "[QRﾀｯﾁﾘｰﾀﾞ]"
            H3_Area1.Visible = True
            '   ( QR Reader 環境 )
            Set QRr_Object = PDQComm1
            Set QRr_MsgObject = QRMsg
            Set QRr_ErrMsgObj = ErrMsg
            Call QRInit
        Case Else
            K_Sykbnm2.Caption = ""
            H1_Area1.Visible = True
    End Select
    ProcHB = mHeadProc
    Call DispChange(ProcHB)
    Call SyoriName(SyoriKB)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_NYmd" Then iH1_NYmd.ShowCalender
    If PB_Look.Tag = "iB1_HKbn" Then Call Look_Kubun: Exit Sub
    If PB_Look.Tag = "iB1_HYmd" Then iB1_HYmd.ShowCalender
    If PB_Look.Tag = "iB1_Sikyu1" Then Call Look_KTorcd(0): Exit Sub
End Sub
'+----------------------------+
'+       名称マスタ検索        +
'+----------------------------+
Private Sub Look_Kubun()

    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "JAMA発注区分"
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
        iB1_HKbn = RV_Code
        Call iB1_HKbn_LostFocus
    End If

End Sub
'+----------------------------+
'+       契約先検索(得意先)    +
'+----------------------------+
Private Sub Look_KTorcd(KTKbn As Integer)
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 5
    
    RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        Select Case KTKbn
            Case 0: iB1_Sikyu1 = RV_TorcdK
                    B1lb_KTornm = RV_TorRName
        End Select
    End If
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
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
'*----------------------------------------------*
'*      納 入 日  ( iH1_NYmd )       *
'*----------------------------------------------*
Private Sub iH1_NYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_NYmd_LostFocus()
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'*----------------------------------------------*
'*      発注区分  ( iB1_HKbn )       *
'*----------------------------------------------*
Private Sub iB1_HKbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_HKbn_LostFocus()
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B1lb_HKbn = ""
    If MeisyoGet("JAMA発注区分", iB1_HKbn) Then
        B1lb_HKbn = SYM.値名称
    End If
End Sub

'*----------------------------------------------*
'*      発 行 日  ( iB1_HYmd )       *
'*----------------------------------------------*
Private Sub iB1_HYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_HYmd_LostFocus()
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'*----------------------------------------------*
'*      発注区分  ( iB1_Sikyu1 )       *
'*----------------------------------------------*
Private Sub iB1_Sikyu1_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Sikyu1_LostFocus()
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B1lb_KTornm = ""
    If TorNmGet(iB1_Sikyu1, 5) Then
        B1lb_KTornm = TRM.略称
    End If
End Sub
'*----------------------------------------------*
'*      出荷場  ( iB1_Syukaba )       *
'*----------------------------------------------*
Private Sub iB1_Syukaba_GotFocus()
    B_TabDummy.Tag = ""
End Sub
'+-----------------------------+
'+  iB1_Hinbn
'+-----------------------------+
Private Sub iB1_Hinbn_GotFocus()
    Set ActObj = iB1_Hinbn
    B_TabDummy.Tag = "S"
End Sub
Private Sub iB1_Hinbn_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Hinbn_LostFocus()
    HNM.表示品番 = ""
    If RTrim(iB1_Hinbn) <> "" Then
        HNM.品番 = RTrim(iB1_Hinbn)
        Call HNMRead("")
    End If
    If RTrim(HNM.表示品番) <> "" Then
        VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 1) = RTrim(HNM.表示品番)
    Else
        VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 1) = iB1_Hinbn
    End If
    VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = iB1_Hinbn
End Sub
'+-----------------------------+
'+  iB1_Seban
'+-----------------------------+
Private Sub iB1_Seban_GotFocus()
    Set ActObj = iB1_Seban
End Sub
Private Sub iB1_Seban_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Seban_LostFocus()
    VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 2) = iB1_Seban
End Sub
'+-----------------------------+
'+  iB1_Syuyo
'+-----------------------------+
Private Sub iB1_Syuyo_GotFocus()
    Set ActObj = iB1_Syuyo
End Sub
Private Sub iB1_Syuyo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Syuyo_LostFocus()
    If CisFun.Val2(iB1_Syuyo) <> 0 Then
        VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 3) = iB1_Syuyo
    Else
        VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 3) = ""
    End If
End Sub
'+-----------------------------+
'+  iB1_Maisu
'+-----------------------------+
Private Sub iB1_Maisu_GotFocus()
    Set ActObj = iB1_Maisu
End Sub
Private Sub iB1_Maisu_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Maisu_LostFocus()
    If CisFun.Val2(iB1_Maisu) <> 0 Then
        VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 4) = iB1_Maisu
        If CisFun.Val2(iB1_Syuyo) <> 0 Then
            gLong = (CisFun.Val2(iB1_Syuyo) * CisFun.Val2(iB1_Maisu))
            iB1_HSuryo = Format(gLong, "####")
        End If
    Else
        VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 4) = ""
    End If
End Sub
'+-----------------------------+
'+  iB1_HSuryo
'+-----------------------------+
Private Sub iB1_HSuryo_GotFocus()
    Set ActObj = iB1_HSuryo
    B_TabDummy.Tag = "E"
End Sub
Private Sub iB1_HSuryo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_HSuryo_LostFocus()
    If CisFun.Val2(iB1_HSuryo) <> 0 Then
        VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 5) = iB1_HSuryo
    Else
        VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 5) = ""
    End If
End Sub
'+-----------------------------+
'+  B_STabDummy
'+-----------------------------+
Private Sub B_STabDummy_GotFocus()
    If B_TabDummy.Tag = "S" Then
        Debug.Print Input_Area.Tag
        If Val(Input_Area.Tag) = 1 Then
            iB1_Syukaba.SetFocus
        Else
            Call TextToVsGrid(Val(Input_Area.Tag))
            Call Grid_RowChange(vbKeyUp)
            iB1_Hinbn.SetFocus
        End If
    Else
        SendKeys "{Tab}"
    End If
End Sub
'+-----------------------------+
'+  B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
    If B_TabDummy.Tag = "E" Then
        Call TextToVsGrid(Val(Input_Area.Tag))
        Call Grid_RowChange(vbKeyDown)
        iB1_Hinbn.SetFocus
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
    Dim H1Mode          As Boolean
    Dim H1Color         As Long
    Dim H2Mode          As Boolean
    Dim H2Color         As Long
    Dim H3Mode          As Boolean
    Dim H3Color         As Long
    Dim H4Mode          As Boolean
    Dim H4Color         As Long
    Dim B1Mode          As Boolean
    Dim B1Color         As Long
    
    H1Mode = False:     H1Color = gPLostSel
    H2Mode = False:     H2Color = gPLostSel
    H3Mode = False:     H3Color = gPLostSel
    H4Mode = False:     H4Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "H2":      H2Mode = True:  H2Color = gPGotSel
        Case "H3":      H3Mode = True:  H3Color = gPGotSel
        Case "H4":      H4Mode = True:  H4Color = gPGotSel
        Case "B1":      B1Mode = True:  B1Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    H3_Area1.Enabled = H3Mode
    H3_Area1.BackColor = H3Color
    H3_Area1.Visible = H3Mode
    
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B1_Area2.Enabled = B1Mode
    B1_Area2.BackColor = B1Color
    
    PB_ADD.Visible = H1Mode Or H3Mode
    PB_MNT.Visible = H1Mode Or H3Mode
    PB_DEL.Visible = H1Mode Or H3Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = B1Mode
    PB_CHG.Visible = Not B1Mode
    If Trim(mSaveHeadProc) = "H1" Then
       PB_CHG.Visible = False
    End If
    
    If Rv_NCall = "" Then
    Else
       PB_END.BtnCaption = "戻る"
    End If
'
    Set QRr_Object = PDQComm1
    Call QR_Comm_Close
    If ProcHB = "H3" Then
       Call QR_Comm_Open
    End If

    Call GridTextEnabled(B1Mode)
' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False

End Function
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   Dim W_Color     As Long
   W_Color = &H80000008 '初期色ｾｯﾄ(黒)
   Select Case SyoriKB
     Case "ADD"
          SyoriNM = "追加"
          K_Sykbnm.ForeColor = &HFFFF&   '(黄色)
          PB_ADD.Visible = False
     Case "MNT"
          SyoriNM = "修正"
          K_Sykbnm.ForeColor = &HFF0000  '(青色)
     Case "DEL"
          SyoriNM = "削除"
          K_Sykbnm.ForeColor = &HFF&     '(赤色)
          W_Color = &HFF&
     Case Else
          SyoriNM = ""
   End Select
   For Each gObj In Me.Controls
       If gObj.Name Like "i*" Then gObj.ForeColor = W_Color
       If gObj.Name Like "*Comb*" Then gObj.ForeColor = W_Color
'      * PF-Key 表示設定 *
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
'**************************************
'*      Head/Body 初期化処理           *
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
       CisVsGrid.Clear
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)

    If ProcHB = "H1" Then
        If Not Head1Chk(mHeadProc) Then GoTo ReturnPress_Ed
        If SyoriKB = "DEL" Then
            If Not DBDelete Then
                ProcHB = mHeadProc
                Call DispChange(ProcHB)
                GoTo ReturnPress_Ed
            End If
            Call HeadBodyClear("B")
            ProcHB = mHeadProc
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
        Else
            ProcHB = "B1"
            Call DispChange(ProcHB)
            VsGrid1.Row = 1
            Call VsGridToText(VsGrid1.Row)
            Set ActObj = iB1_Hinbn
            Input_Area.Tag = "1"
        End If
        GoTo ReturnPress_Ed
    End If

    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        
        
        Call HeadBodyClear("B")
        ProcHB = mHeadProc
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        GoTo ReturnPress_Ed
    End If

ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk(ChkType As String) As Boolean
    Head1Chk = False
    
'( 納入日 )
    If RTrim$(iH1_NYmd) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  納入日を入力して下さい  "
            .MB_Button = Error
            .MBOX
        End With
        iH1_NYmd.SetFocus
        Exit Function
    Else
        If Not iH1_NYmd.cChkResult Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    正しい日付を入力して下さい         "
                 .MB_Button = Error
                 .MBOX
             End With
             iH1_NYmd.SetFocus
             Exit Function
         End If
    End If
    
'( 納入番号 )
    If RTrim$(iH1_Nouno) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  納番を入力して下さい  "
            .MB_Button = Error
            .MBOX
        End With
        iH1_Nouno.SetFocus
        Exit Function
    End If
    
'( 納品書頁 )
    If CisFun.Val2(iH1_Page) = 0 Then iH1_Page = " 1"
    
'( データ確認 )
    
    If Not DBInput(ChkType) Then Exit Function
    
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    Call TextToVsGrid(Val(Input_Area.Tag))
    Call GridTextEnabled(False)
    gCnt1 = VsGrid1.Row
    
'ﾒｯｾｰｼﾞﾎﾞｯｸｽ設定
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With

'納入先
    If RTrim(iB1_Nonyu) = "" Then
        CisFun.MB_MSG(2) = "    納入先を入力して下さい         "
        CisFun.MBOX
        Set ErrObj = iB1_Nonyu
        GoTo Body1Chk_Err
    End If
'受入
    If RTrim(iB1_Ukeir) = "" Then
        CisFun.MB_MSG(2) = "    受入を入力して下さい         "
        CisFun.MBOX
        Set ErrObj = iB1_Ukeir
        GoTo Body1Chk_Err
    End If
'支給元体系
    If RTrim(iB1_SikyuK) = "" Then
        CisFun.MB_MSG(2) = "    支給元（体系）を入力して下さい         "
        CisFun.MBOX
        Set ErrObj = iB1_SikyuK
        GoTo Body1Chk_Err
    End If
'支給元①
    If RTrim(iB1_Sikyu1) = "" Then
        CisFun.MB_MSG(2) = "    支給元（①）を入力して下さい         "
        CisFun.MBOX
        Set ErrObj = iB1_Sikyu1
        GoTo Body1Chk_Err
    End If
'受注者
    If RTrim(iB1_Siire) = "" Then
        CisFun.MB_MSG(2) = "    受注者を入力して下さい         "
        CisFun.MBOX
        Set ErrObj = iB1_Siire
        GoTo Body1Chk_Err
    End If

    Set ErrObj = Nothing
'明細チェック
    gCnt3 = 0
'<< 行を詰める >>
    With VsGrid1
        For gCnt1 = 1 To 10
            If RTrim(.TextMatrix(gCnt1, 1)) = "" Then
                .TextMatrix(gCnt1, 2) = ""
                .TextMatrix(gCnt1, 3) = ""
                .TextMatrix(gCnt1, 4) = ""
                .TextMatrix(gCnt1, 5) = ""
                .TextMatrix(gCnt1, 6) = ""
                For gCnt2 = gCnt1 + 1 To 10
                    If RTrim(.TextMatrix(gCnt2, 1)) <> "" Then
                        .TextMatrix(gCnt1, 1) = .TextMatrix(gCnt2, 1)
                        .TextMatrix(gCnt1, 2) = .TextMatrix(gCnt2, 2)
                        .TextMatrix(gCnt1, 3) = .TextMatrix(gCnt2, 3)
                        .TextMatrix(gCnt1, 4) = .TextMatrix(gCnt2, 4)
                        .TextMatrix(gCnt1, 5) = .TextMatrix(gCnt2, 5)
                        .TextMatrix(gCnt1, 6) = .TextMatrix(gCnt2, 6)
                        .TextMatrix(gCnt2, 1) = ""
                        .TextMatrix(gCnt2, 2) = ""
                        .TextMatrix(gCnt2, 3) = ""
                        .TextMatrix(gCnt2, 4) = ""
                        .TextMatrix(gCnt2, 5) = ""
                        .TextMatrix(gCnt2, 6) = ""
                        Exit For
                    End If
                Next gCnt2
            Else
                gCnt3 = gCnt3 + 1
            End If
        Next gCnt1
        If gCnt3 = 0 Then
            CisFun.MB_MSG(1) = "    品番を入力して下さい！            "
            CisFun.MBOX
            Set ErrObj = iB1_Hinbn
            gCnt1 = 1
            GoTo Body1Chk_Err
        End If
        
        For gCnt1 = 1 To 10
            If RTrim(.TextMatrix(gCnt1, 1)) = "" Then
                Exit For
            End If
        '<< 重複チェック >>
            If RTrim(.TextMatrix(gCnt1, 1)) <> "" Then
                For gCnt2 = gCnt1 + 1 To 10
                    If RTrim(.TextMatrix(gCnt1, 1)) = RTrim(.TextMatrix(gCnt2, 1)) Then
                        CisFun.MB_MSG(1) = "    品番が重複しています。            "
                        CisFun.MB_MSG(3) = "    " & gCnt1 & "行目と " & gCnt2 & "行目            "
                        CisFun.MBOX
                        Set ErrObj = iB1_Hinbn
                        gCnt1 = gCnt2
                        GoTo Body1Chk_Err
                    End If
                Next gCnt2
            End If
        ' 背番号半角チェック
            If RTrim(.TextMatrix(gCnt1, 2)) <> "" And Not HalfSizeChk(.TextMatrix(gCnt1, 2)) Then
                With CisFun
                     .MB_Lines = 4
                     .MB_MSG(2) = "    背番号は半角にて入力して下さい            "
                     .MB_Title = ""
                     .MB_Button = Error
                     .MBOX
                End With
                Set ErrObj = iB1_Seban
                GoTo Body1Chk_Err
            End If
        '<< 発注数チェック >>
            If CisFun.Val2(.TextMatrix(gCnt1, 5)) = 0 Then
                CisFun.MB_MSG(1) = "    " & gCnt1 & "行目【　納入数　必須！】            "
                CisFun.MB_MSG(3) = "    納入数を入力して下さい。            "
                CisFun.MBOX
                Set ErrObj = iB1_Maisu
                GoTo Body1Chk_Err
            End If
            
        '<< 品番マスタチェック >>
            HNM.品番 = RTrim(VsGrid1.TextMatrix(gCnt1, 6))
            If Not HNMRead("", 1) Then
                CisFun.MB_MSG(1) = "    " & gCnt1 & "行目【　品番マスタ　未登録！】            "
                CisFun.MB_MSG(3) = "    品番 = " & RTrim(VsGrid1.TextMatrix(gCnt1, 1))
                CisFun.MBOX
                Set ErrObj = iB1_Hinbn
                GoTo Body1Chk_Err
            End If
        '<< 取引先変換マスタチェック >>
            gSL_Select = "SELECT COUNT(契約先) 件数 FROM 取引先変換マスタ"
            If RTrim(iB1_Sikyu2) = "" Then
                gSL_Select = gSL_Select & " WHERE 支給元 = '" & RTrim(iB1_Sikyu1) & "'"
            Else
                gSL_Select = gSL_Select & " WHERE 支給元 = '" & RTrim(iB1_Sikyu2) & "'"
            End If
            gSL_Select = gSL_Select & "   AND 契約取引先CD = '" & RTrim(iB1_Siire) & RTrim(iB1_SiireKj) & "'"
            gSL_Select = gSL_Select & "   AND (品番 = '' OR 品番 = '" & RTrim(VsGrid1.TextMatrix(gCnt1, 6)) & "') "
            If THMRead(gSL_Select, 1) Then
                If THM.件数 = 0 Then
                    CisFun.MB_MSG(1) = "    " & gCnt1 & "行目【　取引先変換マスタ　未登録！】            "
                    CisFun.MBOX
                    Set ErrObj = iB1_Siire
                    GoTo Body1Chk_Err
                End If
            End If
        '<< 品番取引先マスタチェック >>
            gSL_Select = "SELECT HT.契約先,HT.代行契約先,HT.取引先,HT.受入 FROM 取引先変換マスタ TH"
            gSL_Select = gSL_Select & "  LEFT OUTER JOIN 品番取引先マスタ HT"
            gSL_Select = gSL_Select & "    ON  HT.取引先 = TH.取引先"
            gSL_Select = gSL_Select & "    AND HT.受入 = TH.受入"
            gSL_Select = gSL_Select & "    AND HT.品番 = '" & RTrim(VsGrid1.TextMatrix(gCnt1, 6)) & "' "
            gSL_Select = gSL_Select & "    AND HT.契約先 = TH.契約先"
            gSL_Select = gSL_Select & "    AND HT.代行契約先 = TH.代行契約先"
            gSL_Select = gSL_Select & "    AND HT.品目 = 1"
            If RTrim(iB1_Sikyu2) = "" Then
                gSL_Select = gSL_Select & " WHERE TH.支給元 = '" & RTrim(iB1_Sikyu1) & "'"
            Else
                gSL_Select = gSL_Select & " WHERE TH.支給元 = '" & RTrim(iB1_Sikyu2) & "'"
            End If
            gSL_Select = gSL_Select & "   AND TH.契約取引先CD = '" & RTrim(iB1_Siire) & RTrim(iB1_SiireKj) & "'"
            gSL_Select = gSL_Select & "   AND (TH.品番 = '' OR TH.品番 = '" & RTrim(VsGrid1.TextMatrix(gCnt1, 6)) & "') "
            gSL_Select = gSL_Select & "   AND HT.品番 IS NOT NULL"
            gSL_Select = gSL_Select & " ORDER BY TH.取引先,TH.受入,TH.契約先,TH.代行契約先"
            If Not HTMRead(gSL_Select, 1) Then
                CisFun.MB_MSG(1) = "    " & gCnt1 & "行目【　品番取引先マスタ　未登録！】            "
                CisFun.MBOX
                Set ErrObj = iB1_Hinbn
                GoTo Body1Chk_Err
            End If
            B1lb_JUkeir = RTrim(HTM.受入)
            B1lb_JSiire = RTrim(HTM.取引先)
            B1lb_JKTorcd = RTrim(HTM.契約先)
            B1lb_JDaikoK = RTrim(HTM.代行契約先)
        Next gCnt1

    End With

    
    Body1Chk = True
    Exit Function
Body1Chk_Err:
    VsGrid1.Row = gCnt1
    Call VsGridToText(VsGrid1.Row)
    Call GridTextEnabled(True)
    ErrObj.SetFocus
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    
    Call HeadBodyClear("B")
    
    gSL_Select = "SELECT MK.仕入先,MK.受入,MK.契約先,MK.代行契約先,MK.支給元,MK.納入日,MK.納入便,"
    gSL_Select = gSL_Select & "   MK.納入番号,MK.納品書頁,MK.納品書行,MK.品番,背番号,枚数,収容数,検収数,"
    gSL_Select = gSL_Select & "   MK.支給元体系,MK.支給元1,MK.支給元2,MK.出荷場,MK.受注者,MK.受注者工場,MK.発注者,MK.発注者工場,"
    gSL_Select = gSL_Select & "   MK.契約先受入,MK.発行日,MK.発注区分,MK.カード,MK.機番,MK.総頁,"
    gSL_Select = gSL_Select & " ISNULL(HN.表示品番,'') 表示品番,ISNULL(SY.値名称,'') 発注区分名,"
    gSL_Select = gSL_Select & " ISNULL(TR.略称,'') 契約先名"
    gSL_Select = gSL_Select & " FROM 受給検収テーブル MK"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "   ON HN.品番 = MK.品番"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON TR.取引先CD = MK.支給元1"
    gSL_Select = gSL_Select & "   AND TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR.契約先区分 = 1"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY"
    gSL_Select = gSL_Select & "   ON SY.区分名称 = 'JAMA発注区分'"
    gSL_Select = gSL_Select & "   AND SY.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY.値 = MK.発注区分"
    gSL_Select = gSL_Select & " WHERE 納入番号 = '" & RTrim(iH1_Nouno) & "'"
    gSL_Select = gSL_Select & "   AND 納品書頁 = " & CisFun.Val2(iH1_Page)
    gSL_Select = gSL_Select & "   AND 納入日 = '" & RTrim(iH1_NYmd) & "'"
    gSL_Select = gSL_Select & " UNION "
    gSL_Select = gSL_Select & " SELECT MK.仕入先,MK.仕入先受入 受入,MK.契約先,MK.代行契約先,'' 支給元,MK.納入日,MK.納入便, "
    gSL_Select = gSL_Select & "   MK.納入番号,MK.納品書頁,MK.納品書行,MK.品番,MK.背番号,MK.枚数,MK.収容数,MK.納入数 検収数,"
    gSL_Select = gSL_Select & "   MK.支給元体系,MK.支給元1,MK.支給元2,MK.出荷場,MK.受注者,MK.受注者工場,MK.発注者,MK.発注者工場,"
    gSL_Select = gSL_Select & "   MK.受入 契約先受入, MK.発行日, MK.発注区分, MK.カード, MK.機番, MK.総頁,"
    gSL_Select = gSL_Select & " ISNULL(HN.表示品番,'') 表示品番,ISNULL(SY.値名称,'') 発注区分名,"
    gSL_Select = gSL_Select & " ISNULL(TR.略称,'') 契約先名"
    gSL_Select = gSL_Select & " FROM 前工程入荷ワーク MK"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "   ON HN.品番 = MK.品番"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON TR.取引先CD = MK.支給元1"
    gSL_Select = gSL_Select & "   AND TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR.契約先区分 = 1"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY"
    gSL_Select = gSL_Select & "   ON SY.区分名称 = 'JAMA発注区分'"
    gSL_Select = gSL_Select & "   AND SY.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY.値 = MK.発注区分"
    gSL_Select = gSL_Select & " WHERE 処理端末 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND 納入番号 = '" & RTrim(iH1_Nouno) & "'"
    gSL_Select = gSL_Select & "   AND 納品書頁 = " & CisFun.Val2(iH1_Page)
    gSL_Select = gSL_Select & "   AND 納入日 = '" & RTrim(iH1_NYmd) & "'"
    
    On Error GoTo DBInput_Err:
    If Not JYKRead(gSL_Select, 0) Then
        Call JYKClose
        If SyoriKB <> "ADD" Then
            If InputKey = "H3" Then
                QRDataArea.エラー内容 = "該当データなし！"
            Else
                If InputKey <> "INV" Then
                    With CisFun
                        .MB_Lines = 5
                        .MB_MSG(2) = "  該当データがありません  "
                        .MB_Button = Error
                        .MBOX
                    End With
                End If
            End If
            GoTo DBInput_Ed
        End If
        If SyoriKB = "ADD" And InputKey = "H3" Then
            Call BodySetQR
        End If
        
    Else
        Call BodySet(InputKey)
        
        If SyoriKB = "ADD" Then
            If InputKey = "H3" Then
                QRDataArea.エラー内容 = "登録済み伝票です！"
            Else
                If InputKey <> "INV" Then
                    With CisFun
                        .MB_Lines = 5
                        .MB_MSG(2) = "  　　登録済み伝票です！　　  "
                        .MB_Button = Error
                        .MBOX
                    End With
                End If
            End If
            GoTo DBInput_Ed
        End If
        If B1_Sosin.Visible Then
            If InputKey = "H3" Then
                QRDataArea.エラー内容 = "TNS送信済みデータ"
                GoTo DBInput_Ed
            Else
                If InputKey <> "INV" Then
                    With CisFun
                        .MB_Lines = 5
                        .MB_MSG(2) = "  　　TNS送信済みデータですがよろしいですか？　　  "
                        .MB_Button = OK_CAN
                        If Not .MBOX Then
                            GoTo DBInput_Ed
                        End If
                    End With
                End If
            End If
        End If
        
    End If
    
    DBInput = True
    
DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'**************************************
'*          BODY 内容ｾｯﾄ処理           *
'**************************************
Private Sub BodySet(SetType As String)
'
    With JYK
        iB1_HKbn = .発注区分
        B1lb_HKbn = .発注区分名
        iB1_HYmd = .発行日
        iB1_TPage = Format(.総頁, "##")
        
        iB1_Nonyu = .発注者
        iB1_NonyuKj = .発注者工場
        iB1_Ukeir = .契約先受入
        B1lb_JUkeir = .受入
        iB1_NBin = Format(.納入便, "##")
        iB1_Kiban = .機番
        iB1_Card = .カード
        iB1_SikyuK = .支給元体系
        iB1_Sikyu1 = .支給元1
        B1lb_KTornm = .契約先名
        B1lb_JKTorcd = .契約先
        B1lb_JDaikoK = .代行契約先
        iB1_Sikyu2 = .支給元2
        iB1_Siire = .受注者
        iB1_SiireKj = .受注者工場
        B1lb_JSiire = .仕入先
        iB1_Syukaba = .出荷場
        
        If RTrim(.TNS作成日) <> "" Then
            With CisFun
                .DateE_BefVal = JYK.TNS作成日
                .DateEdit
                B1lb_TNSYmd = .DateE_AfVal
            End With
            B1_Sosin.Visible = True
        Else
            B1_Sosin.Visible = False
        End If
    End With
    
    Call GridSet(SetType)
    
End Sub
Private Sub BodySetQR()
'
    With typQRAfter
        iB1_HKbn = .発注区分
        Call iB1_HKbn_LostFocus
        iB1_HYmd = .発行日
        iB1_TPage = Format(.総ページ数, "##")
        
        iB1_Nonyu = .納入先
        iB1_NonyuKj = .納入先工区
        iB1_Ukeir = .受入
        iB1_NBin = .納入便
        iB1_Kiban = .機番
        iB1_Card = .カード丼
        iB1_SikyuK = .支給元体系
        iB1_Sikyu1 = .支給元1
        Call iB1_Sikyu1_LostFocus
        iB1_Sikyu2 = .支給元2
        iB1_Siire = .受注者
        iB1_SiireKj = .受注者事業所
        iB1_Syukaba = .出荷場
            
        B1_Sosin.Visible = False
        
        Call CisVsGrid.Clear
        With VsGrid1
            .Redraw = False
            For gInt = 1 To 10
                If RTrim(typQRAfter.品番(gInt)) <> "" Then
                    .TextMatrix(gInt, 0) = gInt
                    .TextMatrix(gInt, 1) = typQRAfter.品番(gInt)
                    .TextMatrix(gInt, 2) = typQRAfter.背番号(gInt)
                    .TextMatrix(gInt, 3) = typQRAfter.収容数(gInt)
                    .TextMatrix(gInt, 4) = typQRAfter.箱数(gInt)
                    .TextMatrix(gInt, 5) = typQRAfter.納入数(gInt)
                    .TextMatrix(gInt, 6) = typQRAfter.品番(gInt)
                End If
            Next

            .Row = 1: .Col = 1
            .Redraw = True
        End With
    End With
    
    
End Sub
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+      削 除 処 理      +
'+----------------------+
Private Function DBDelete()
    DBDelete = False
'  ( 確認メッセージ )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     削除処理を行います      "
        .MB_Title = "削除処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    With CisDB
        On Error GoTo DBDelete_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 削除処理 )
    ' 発注画面入力ワーク
        .SQL = "DELETE FROM 受給検収テーブル "
        .SQL = .SQL & " WHERE 納入番号 = '" & RTrim(iH1_Nouno) & "'"
        .SQL = .SQL & "   AND 納品書頁 = " & CisFun.Val2(iH1_Page) & ""
        .SQL = .SQL & "   AND 納入日 = '" & RTrim(iH1_NYmd) & "'"
        .DBExec
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

        iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")

        On Error GoTo 0
    End With
    DBDelete = True
    Exit Function
DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'*-------------------------------------------------*
'*      デ ー タ 更 新  確 認 メ ッ セ ー ジ
'*-------------------------------------------------*
Private Function CheckMessage() As Boolean
    CheckMessage = False
    
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     前工程発行JAMA 入荷実績処理を行います      "
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Title = "実績更新処理"
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Call VsGridToText(Val(Input_Area.Tag))
            Call GridTextEnabled(True)
            Exit Function
        End If
    End With
    
    CheckMessage = True
    
End Function
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False
    Static AddFlg               As Boolean
    Static SaveKanriNo          As Long
    
'   ( 確認メッセージ )
    If Not CheckMessage Then Exit Function
    
    CisDB.ConnectNo = 0
    Call CisDB.DBTran(TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
    If SyoriKB = "MNT" Then
        gSL_Select = "DELETE FROM 受給検収テーブル"
        gSL_Select = gSL_Select & " WHERE 納入番号 = '" & RTrim(iH1_Nouno) & "'"
        gSL_Select = gSL_Select & "   AND 納品書頁 = " & CisFun.Val2(iH1_Page)
        gSL_Select = gSL_Select & "   AND 納入日 = '" & RTrim(iH1_NYmd) & "'"
        CisDB.SQL = gSL_Select
        CisDB.DBExec
    End If
'   検収管理NO獲得
        With CisDB
            .SQL = "番号獲得"
            .StoadoCount = 3
            .ParaNo = 0: .ParaIO = Return用
            .ParaNo = 1: .ParaIO = Input用
            .ParaNo = 2: .ParaIO = OutPut用
            .ParaNo = 3: .ParaIO = Return用

            .ParaNo = 1: .ParaValue = 19

            If Not .DBStored Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "     受給検収管理NOの採番に失敗しました    "
                    .MB_Title = "採番処理"
                    .MB_Button = Error
                    If Not .MBOX Then GoTo DBPut_Err
                End With
            Else
                .ParaNo = 2: SaveKanriNo = .ParaValue
            End If
        End With

'   マスタ項目移送
    For gCnt1 = 1 To VsGrid1.Rows - 1
        If RTrim(VsGrid1.TextMatrix(gCnt1, 1)) <> "" Then
            Call ItemsClearJYK
            With JYK
                .受給管理NO = SaveKanriNo
                .受給管理行NO = gCnt1
                .検収日 = iH1_NYmd
                .仕入先 = RTrim(B1lb_JSiire)
                .受入 = RTrim(B1lb_JUkeir)
                .契約先 = RTrim(B1lb_JKTorcd)
                .代行契約先 = RTrim(B1lb_JDaikoK)
                If RTrim(iB1_Sikyu2) = "" Then
                    .支給元 = RTrim(iB1_Sikyu1)
                Else
                    .支給元 = RTrim(iB1_Sikyu2)
                End If
                .納入日 = iH1_NYmd
                .納入便 = Format(iB1_NBin, "00")
                .納入番号 = RTrim(iH1_Nouno)
                .納品書頁 = CisFun.Val2(iH1_Page)
                
                .納品書行 = gCnt1
                .品番 = VsGrid1.TextMatrix(gCnt1, 6)
                .背番号 = VsGrid1.TextMatrix(gCnt1, 2)
                .収容数 = CisFun.Val2(VsGrid1.TextMatrix(gCnt1, 3))
                .枚数 = CisFun.Val2(VsGrid1.TextMatrix(gCnt1, 4))
                .検収数 = CisFun.Val2(VsGrid1.TextMatrix(gCnt1, 5))
                
                .支給元体系 = RTrim(iB1_SikyuK)
                .支給元1 = RTrim(iB1_Sikyu1)
                .支給元2 = RTrim(iB1_Sikyu2)
                .出荷場 = RTrim(iB1_Syukaba)
                .契約先仕入先 = RTrim(iB1_Siire) & RTrim(iB1_SiireKj)
                .契約先納入先 = RTrim(iB1_Nonyu) & RTrim(iB1_NonyuKj)
                .契約先受入 = RTrim(iB1_Ukeir)
                
                .受注者 = RTrim(iB1_Siire)
                .受注者工場 = RTrim(iB1_SiireKj)
                .発注者 = RTrim(iB1_Nonyu)
                .発注者工場 = RTrim(iB1_NonyuKj)
                .発行日 = iB1_HYmd
                .発注区分 = iB1_HKbn
                .カード = iB1_Card
                .機番 = iB1_Kiban
                .総頁 = CisFun.Val2(iB1_TPage)
                            
                .伝票種類 = 11
                .作成種別 = "M"
                    
            End With
    
            JYK.作成者 = gTanto
            Call JYKInsert
        End If
    Next gCnt1
    
    CisDB.ConnectNo = 0
    Call CisDB.DBTran(TransCommit)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    If SyoriKB = "ADD" Then
        iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "#,###")
    Else
        iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "#,###")
    End If
    
    
    On Error GoTo 0
    
    DBPut = True
    Exit Function
    
DBPut_Err:
'''    Call JYKClose(1)
    
    CisDB.ConnectNo = 0
    Call CisDB.DBTran(TransRollback) ' ﾄﾗﾝｻﾞｸｼｮﾝ破棄
    CisFun.ErrorBox
    End
End Function
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理　関　連
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 0, 0, "№")
        Call .FixedSet(中中, 左中, あり, 0, 1, "品　番")
        Call .FixedSet(中中, 左中, あり, 0, 2, "背番号")
        Call .FixedSet(中中, 右中, あり, 0, 3, "収容数")
        Call .FixedSet(中中, 右中, あり, 0, 4, "箱 数")
        Call .FixedSet(中中, 右中, あり, 0, 5, "納入数")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 中中, あり, 4, 6, "S品番")
'        Call .FixedSet(中中, 左中, なし, 0,9,"Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet(SetType As String)
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGrid1
       .Redraw = False
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not JYK_RDSTS
        With VsGrid1
            ID = ID + 1
            If JYK.納品書行 > 10 Then Exit Do
            'No
            .TextMatrix(JYK.納品書行, 0) = JYK.納品書行
            '品番
            If RTrim(JYK.表示品番) <> "" Then
                .TextMatrix(JYK.納品書行, 1) = JYK.表示品番
            Else
                .TextMatrix(JYK.納品書行, 1) = JYK.品番
            End If
            ' 背番号
            .TextMatrix(JYK.納品書行, 2) = JYK.背番号
            ' 収容数
            .TextMatrix(JYK.納品書行, 3) = Format(JYK.収容数, "#,###")
            ' 枚数
            .TextMatrix(JYK.納品書行, 4) = Format(JYK.枚数, "#,###")
            ' 発注数
            .TextMatrix(JYK.納品書行, 5) = Format(JYK.検収数, "#,###")
            '品番
            .TextMatrix(JYK.納品書行, 6) = JYK.品番
        End With
        Call JYKReadNext
    Loop
    Call JYKClose
    With VsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    
End Sub
'==========================================================='
'   入力テキスト表示・非表示
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
    Input_Area.Visible = EnabledFlg
End Sub
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown Then Exit Sub

'テキストからグリッドに退避
    Call TextToVsGrid(Val(Input_Area.Tag))
    gInt = VsGrid1.Row
    '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    '↓
    If KeyCode = vbKeyDown Then
        If gInt + 1 > VsGrid1.Rows - 1 Then GoTo Grid_RowChange_End
        gInt = gInt + 1
    End If
    
    VsGrid1.Row = gInt
'グリッドからテキストに表示
    Call VsGridToText(VsGrid1.Row)
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    With VsGrid1
        .TextMatrix(tRow, 6) = RTrim(iB1_Hinbn)
        HNM.表示品番 = ""
        If RTrim(iB1_Hinbn) <> "" Then
            HNM.品番 = RTrim(iB1_Hinbn)
            Call HNMRead("")
        End If
        If RTrim(HNM.表示品番) <> "" Then
            VsGrid1.TextMatrix(tRow, 1) = RTrim(HNM.表示品番)
        Else
            VsGrid1.TextMatrix(tRow, 1) = iB1_Hinbn
        End If
        
        .TextMatrix(tRow, 2) = RTrim(iB1_Seban)
        If CisFun.Val2(iB1_Syuyo) <> 0 Then
            .TextMatrix(tRow, 3) = Format(CisFun.Val2(iB1_Syuyo), "#,###")
        End If
        If CisFun.Val2(iB1_Maisu) <> 0 Then
            .TextMatrix(tRow, 4) = Format(CisFun.Val2(iB1_Maisu), "#,###")
        Else
            .TextMatrix(tRow, 4) = ""
        End If
        If CisFun.Val2(iB1_HSuryo) <> 0 Then
            .TextMatrix(tRow, 5) = Format(CisFun.Val2(iB1_HSuryo), "#,###")
        Else
            .TextMatrix(tRow, 5) = ""
        End If
        
    End With
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With VsGrid1
        Input_Area.Tag = tRow
        Input_Area.Move Input_Area.Left, VsGrid1.CellTop, Input_Area.Width, VsGrid1.CellHeight
        '品番
        iB1_Hinbn = RTrim(.TextMatrix(tRow, 6))
        '背番号
        iB1_Seban = RTrim(.TextMatrix(tRow, 2))
        '収容数
        iB1_Syuyo = Format(CisFun.Val2(.TextMatrix(tRow, 3)), "#,###")
        '枚数
        iB1_Maisu = CisFun.Val2(.TextMatrix(tRow, 4))
        '発注数
        iB1_HSuryo = CisFun.Val2(.TextMatrix(tRow, 5))
    End With
End Sub
Private Sub VSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown Then Exit Sub
    Call TextToVsGrid(Val(Input_Area.Tag))
    Call VsGridToText(VsGrid1.Row)
End Sub
Private Sub VSGrid1_Click()
    Call TextToVsGrid(Val(Input_Area.Tag))
'    If VsGrid1.TextMatrix(VsGrid1.Row, 0) = "" Then Exit Sub
    Call GridTextEnabled(True)
    Call VsGridToText(VsGrid1.Row)
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub
'*=======================================================================================================*
'*=======================================================================================================*
'*
'*                      Ｑ  Ｒ  リ  ー  ダ  (　Ｃｏｍｍ　)  処  理
'*
'*=======================================================================================================*
'*=======================================================================================================*
'+--------------------------+
'+      読  取  処  理      +
'+--------------------------+
Private Sub PDQComm1_OnComm()
    
    gStr = QR_CommEvent
    If gStr = "ER" Then GoTo PDQComm1_OnComm_ED
    If gStr <> "OK" Then Exit Sub
    
    With QRDataArea
        If .ステータス <> 0 Then
           Call QRReadMsg("E")
           Exit Sub
        End If
    ' JAMA伝票の納品書・支給書のみ有効
        If .バーコード種類 <> "J" Then
            .エラー内容 = "JAMA伝票のみ対象"
            Call QRReadMsg("E")
            Exit Sub
        End If
        
        If typQRAfter.カード丼 <> "7" And typQRAfter.カード丼 <> "8" Then
           .エラー内容 = "カードが対象外です"
           Call QRReadMsg("E")
           Exit Sub
        End If
        QRr_ReadData = ""
        
        iH1_NYmd = typQRAfter.納入日
        iH1_Nouno = typQRAfter.納入番号
        iH1_Page = typQRAfter.ページNO
        If Not Head1Chk("H3") Then
            Call QRReadMsg("E")
            ProcHB = "H3"
'            Exit Sub
        Else
            ProcHB = "H1"
            Call PB_ENT_Click
'            Exit Sub
        End If
    End With
    
'    ProcHB = "B1"
'    Call DispChange(ProcHB)
    Exit Sub
PDQComm1_OnComm_ED:
    Call HeadBodyClear("B")
    Call QRReadMsg("E")
End Sub
'+---------------------------------------+
'+      名称マスタ初期設定
'+---------------------------------------+
Private Sub MeisyoInit()
' JAMA発注区分
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "JAMA発注区分"
        .区分桁数 = 1
        .値 = "1"
        .制御区分 = "X"
        .値名称 = "号口"
        .作成者 = gTanto
        .検索名 = "JAMA発注区分 検索"
        .コード名 = "発区"
        .名称名 = "JAMA発注区分名"
    End With
    If Not SYMInitCreate(True) Then
        With SYMCreate
            .値 = "2"
            .値名称 = "特配"
        End With
        Call SYMInitCreate
        With SYMCreate
            .値 = "3"
            .値名称 = "臨時"
        End With
        Call SYMInitCreate
        With SYMCreate
            .値 = "4"
            .値名称 = "スポット"
        End With
        Call SYMInitCreate
        With SYMCreate
            .値 = "5"
            .値名称 = "号試"
        End With
        Call SYMInitCreate
    End If
End Sub
