VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CKNZ0210 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "材料発注かんばん個別発行"
   ClientHeight    =   10905
   ClientLeft      =   210
   ClientTop       =   525
   ClientWidth     =   14880
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
   ScaleHeight     =   10905
   ScaleWidth      =   14880
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'なし
      Height          =   495
      Index           =   1
      Left            =   10950
      ScaleHeight     =   495
      ScaleWidth      =   2520
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5745
      Width           =   2520
      Begin Cis3D_v60.CIS3D UC_3D6 
         Height          =   435
         Left            =   30
         Top             =   30
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   767
         ForeColor       =   0
         Caption         =   " 発行総枚数"
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
         Begin Cis3D_v60.CIS3D iT1_Addsu 
            Height          =   315
            Left            =   1350
            Top             =   60
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BackColor       =   16761024
            ForeColor       =   0
            Caption         =   "ZZ,ZZ9"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
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
   Begin Cis3D_v60.CIS3D B1_Area2 
      Height          =   825
      Left            =   1230
      Top             =   2925
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   1455
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
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   735
         Left            =   75
         Top             =   45
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1296
         ForeColor       =   0
         Caption         =   "サイクル"
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
         cAlingnment     =   6
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_Cycle 
            Height          =   330
            Left            =   90
            Top             =   330
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "X-XX-XX"
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D12 
         Height          =   735
         Left            =   1290
         Top             =   45
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   1296
         ForeColor       =   0
         Caption         =   "材  料  名"
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
         cAlingnment     =   6
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_Zainm 
            Height          =   330
            Left            =   90
            Top             =   330
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   735
         Left            =   9810
         Top             =   45
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   1296
         ForeColor       =   0
         Caption         =   "契約先"
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
         cAlingnment     =   6
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_KTorcd 
            Height          =   330
            Left            =   90
            Top             =   330
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "XXXXXXX"
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D23 
         Height          =   720
         Left            =   10995
         Top             =   45
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   1270
         ForeColor       =   0
         Caption         =   "手配区"
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
         cAlingnment     =   6
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_TKbn 
            Height          =   330
            Left            =   240
            Top             =   330
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "XXXX"
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D Back_Torcd 
         Height          =   735
         Left            =   6390
         Top             =   45
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   1296
         ForeColor       =   0
         Caption         =   "メーカー"
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
         cAlingnment     =   6
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_Maker 
            Height          =   330
            Left            =   90
            Top             =   330
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "XXXXXXX:XXXXXXXXX1XXXXX"
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B1lb_MUke 
            Height          =   330
            Left            =   2940
            Top             =   330
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "XX"
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   315
            Index           =   9
            Left            =   3000
            Top             =   15
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            ForeColor       =   0
            Caption         =   "受"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   2
            cPositionY      =   30
         End
      End
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
      Height          =   930
      Left            =   1215
      ScaleHeight     =   870
      ScaleWidth      =   12225
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1815
      Width           =   12285
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   780
         Left            =   60
         Top             =   60
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1376
         ForeColor       =   192
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Seban 
            Height          =   360
            Left            =   75
            TabIndex        =   0
            Top             =   330
            Width           =   1140
            _ExtentX        =   2011
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
            Text            =   "XXXXXXXX"
            MaxLength       =   8
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D BLB_HINBN 
         Height          =   780
         Left            =   5655
         Top             =   60
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "材  質"
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
         Begin CisText_V60.CisText iH1_Zai 
            Height          =   360
            Left            =   90
            TabIndex        =   3
            Top             =   330
            Width           =   3195
            _ExtentX        =   5636
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   780
         Left            =   1365
         Top             =   60
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "仕 入 先"
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
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   60
            TabIndex        =   1
            Top             =   330
            Width           =   990
            _ExtentX        =   1746
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
            Text            =   "XXXXXXX"
            MaxLength       =   7
         End
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1035
            Top             =   315
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   635
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
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D15 
         Height          =   780
         Left            =   4950
         Top             =   60
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   1376
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
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   360
            Left            =   150
            TabIndex        =   2
            Top             =   330
            Width           =   390
            _ExtentX        =   688
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
            Text            =   "XX"
            MaxLength       =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   780
         Left            =   9000
         Top             =   60
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "板厚"
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
         Begin CisText_V60.CisText iH1_Ita 
            Height          =   360
            Left            =   45
            TabIndex        =   4
            Top             =   330
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   635
            cFormat         =   "###0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
            cILength        =   3
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   780
         Left            =   10050
         Top             =   60
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "幅"
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
         Begin CisText_V60.CisText iH1_Haba 
            Height          =   360
            Left            =   45
            TabIndex        =   5
            Top             =   330
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   635
            cFormat         =   "###0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
            cILength        =   3
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   780
         Left            =   11100
         Top             =   60
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "長さ"
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
         Begin CisText_V60.CisText iH1_Nagasa 
            Height          =   360
            Left            =   45
            TabIndex        =   6
            Top             =   330
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   635
            cFormat         =   "###0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
            cILength        =   3
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
      Height          =   720
      Left            =   7350
      ScaleHeight     =   660
      ScaleWidth      =   7710
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   480
      Width           =   7770
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   5160
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
         Left            =   6000
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
         Left            =   6840
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
         Left            =   3465
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
      Begin CisBtn_60.CisBtn PB_NEXT 
         Height          =   585
         Left            =   2610
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
         Left            =   1770
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
      Begin CisBtn_60.CisBtn PB_Prt 
         Height          =   585
         Left            =   4305
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
         BtnType         =   13
         BtnCaption      =   "印刷"
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
      Begin CisBtn_60.CisBtn PB_Del 
         Height          =   585
         Left            =   915
         Top             =   30
         Visible         =   0   'False
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
      Begin CisBtn_60.CisBtn PB_Add 
         Height          =   585
         Left            =   45
         Top             =   30
         Visible         =   0   'False
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
         BtnCaption      =   "追加"
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
      Left            =   960
      TabIndex        =   15
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
      BackColor       =   8438015
      ForeColor       =   0
      Caption         =   "【 材料・発注かんばん個別発行(ﾀｰﾝｱﾗｳﾝﾄﾞ)  】"
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
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14310
         Top             =   75
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   450
         BackColor       =   8438015
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
         cBoderColor1    =   8438015
         cBoderColor2    =   8438015
         cFont3DColor1   =   8454143
         cFont3DColor2   =   16777215
         cFont3D         =   1
         cAlingnment     =   7
         cBoderStyle     =   2
      End
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   285
         Left            =   11805
         Top             =   60
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   503
         BackColor       =   8438015
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
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   65535
         Caption         =   "追加"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   1665
      Left            =   1215
      ScaleHeight     =   1605
      ScaleWidth      =   12225
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3810
      Width           =   12285
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   1470
         Left            =   60
         Top             =   60
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2593
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderStyle     =   1
         Begin Cis3D_v60.CIS3D CIS3D28 
            Height          =   1050
            Left            =   30
            Top             =   390
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   1852
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   60
            Begin Cis3D_v60.CIS3D B1lb_SHYmd 
               Height          =   420
               Left            =   165
               Top             =   300
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   741
               BackColor       =   16777152
               Caption         =   "9999/99/99"
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cBoderStyle     =   1
            End
            Begin Cis3D_v60.CIS3D B1lb_SHBin 
               Height          =   420
               Left            =   1575
               Top             =   300
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   741
               BackColor       =   16777152
               Caption         =   "99"
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cBoderStyle     =   1
            End
         End
         Begin Cis3D_v60.CIS3D BLB_Sur 
            Height          =   525
            Index           =   1
            Left            =   2100
            Top             =   390
            Width           =   8805
            _ExtentX        =   15531
            _ExtentY        =   926
            ForeColor       =   16711680
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   510
               Index           =   8
               Left            =   4095
               Top             =   0
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   900
               ForeColor       =   0
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin Cis3D_v60.CIS3D B1lb_HZSu 
                  Height          =   330
                  Index           =   0
                  Left            =   105
                  Top             =   75
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "###"
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -350
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   510
               Index           =   5
               Left            =   2250
               Top             =   0
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   900
               ForeColor       =   0
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                  Height          =   330
                  Index           =   0
                  Left            =   180
                  Top             =   75
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "##,###.000"
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -100
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   510
               Index           =   6
               Left            =   555
               Top             =   0
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   900
               ForeColor       =   0
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin Cis3D_v60.CIS3D B1lb_Syuki 
                  Height          =   330
                  Index           =   0
                  Left            =   135
                  Top             =   75
                  Width           =   1320
                  _ExtentX        =   2328
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "XXXXXXXXX1"
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cAlingnment     =   1
                  cBoderStyle     =   1
                  cPositionX      =   60
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   510
               Index           =   7
               Left            =   5610
               Top             =   0
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   900
               ForeColor       =   16711680
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin CisText_V60.CisText iB1_HSu 
                  Height          =   360
                  Index           =   0
                  Left            =   345
                  TabIndex        =   7
                  Top             =   75
                  Width           =   525
                  _ExtentX        =   926
                  _ExtentY        =   635
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
                  Text            =   "###"
                  MaxLength       =   3
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
            End
            Begin Cis3D_v60.CIS3D Back_SSu 
               Height          =   510
               Index           =   0
               Left            =   6870
               Top             =   0
               Width           =   1920
               _ExtentX        =   3387
               _ExtentY        =   900
               ForeColor       =   16711680
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin CisText_V60.CisText iB1_SSSu 
                  Height          =   360
                  Index           =   0
                  Left            =   210
                  TabIndex        =   8
                  Top             =   75
                  Width           =   540
                  _ExtentX        =   953
                  _ExtentY        =   635
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
                  Text            =   "###"
                  MaxLength       =   3
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
               Begin CisText_V60.CisText iB1_ESSu 
                  Height          =   360
                  Index           =   0
                  Left            =   1185
                  TabIndex        =   9
                  Top             =   75
                  Width           =   540
                  _ExtentX        =   953
                  _ExtentY        =   635
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
                  Text            =   "###"
                  MaxLength       =   3
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
               Begin Cis3D_v60.CIS3D CIS3D17 
                  Height          =   255
                  Left            =   870
                  Top             =   120
                  Width           =   210
                  _ExtentX        =   370
                  _ExtentY        =   450
                  Caption         =   "～"
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cBoderStyle     =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   510
               Left            =   0
               Top             =   0
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   900
               BackColor       =   12648384
               Caption         =   "1"
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
               cAlingnment     =   7
            End
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   1395
            Index           =   0
            Left            =   10905
            Top             =   45
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   2461
            ForeColor       =   0
            Caption         =   "最終№"
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
            Begin Cis3D_v60.CIS3D B1lb_EndSu 
               Height          =   615
               Left            =   120
               Top             =   510
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   1085
               BackColor       =   16777152
               Caption         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   18
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
               cAlingnment     =   7
               cBoderStyle     =   1
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   345
            Left            =   30
            Top             =   45
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   609
            Caption         =   "収容変更日･便"
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
            cAlingnment     =   6
            cPositionY      =   30
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   345
            Index           =   2
            Left            =   4350
            Top             =   45
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   609
            ForeColor       =   0
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   30
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   345
            Index           =   50
            Left            =   2100
            Top             =   45
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   609
            ForeColor       =   0
            Caption         =   "収容器"
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
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   345
            Index           =   1
            Left            =   7710
            Top             =   45
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   609
            ForeColor       =   16711680
            Caption         =   "追加枚数"
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
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   345
            Index           =   4
            Left            =   8970
            Top             =   45
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   609
            ForeColor       =   16711680
            Caption         =   "再発行№"
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
         Begin Cis3D_v60.CIS3D BLB_Sur 
            Height          =   525
            Index           =   0
            Left            =   2100
            Top             =   915
            Width           =   8790
            _ExtentX        =   15505
            _ExtentY        =   926
            ForeColor       =   16711680
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   510
               Index           =   12
               Left            =   4095
               Top             =   15
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   900
               ForeColor       =   0
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin Cis3D_v60.CIS3D B1lb_HZSu 
                  Height          =   330
                  Index           =   1
                  Left            =   105
                  Top             =   75
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "###"
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -350
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   510
               Index           =   10
               Left            =   2250
               Top             =   15
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   900
               ForeColor       =   0
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                  Height          =   330
                  Index           =   1
                  Left            =   180
                  Top             =   75
                  Width           =   1470
                  _ExtentX        =   2593
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "##,###.000"
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -100
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   510
               Index           =   11
               Left            =   540
               Top             =   15
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   900
               ForeColor       =   0
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin Cis3D_v60.CIS3D B1lb_Syuki 
                  Height          =   330
                  Index           =   1
                  Left            =   150
                  Top             =   90
                  Width           =   1320
                  _ExtentX        =   2328
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "XXXXXXXXX1"
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cAlingnment     =   1
                  cBoderStyle     =   1
                  cPositionX      =   60
               End
            End
            Begin Cis3D_v60.CIS3D Back_HSu 
               Height          =   510
               Index           =   1
               Left            =   5610
               Top             =   15
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   900
               ForeColor       =   16711680
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin CisText_V60.CisText iB1_HSu 
                  Height          =   360
                  Index           =   1
                  Left            =   345
                  TabIndex        =   10
                  Top             =   75
                  Width           =   525
                  _ExtentX        =   926
                  _ExtentY        =   635
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
                  Text            =   "###"
                  MaxLength       =   3
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
            End
            Begin Cis3D_v60.CIS3D Back_SSu 
               Height          =   510
               Index           =   1
               Left            =   6870
               Top             =   15
               Width           =   1920
               _ExtentX        =   3387
               _ExtentY        =   900
               ForeColor       =   16711680
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin CisText_V60.CisText iB1_SSSu 
                  Height          =   360
                  Index           =   1
                  Left            =   210
                  TabIndex        =   11
                  Top             =   75
                  Width           =   540
                  _ExtentX        =   953
                  _ExtentY        =   635
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
                  Text            =   "###"
                  MaxLength       =   3
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
               Begin CisText_V60.CisText iB1_ESSu 
                  Height          =   360
                  Index           =   1
                  Left            =   1185
                  TabIndex        =   12
                  Top             =   75
                  Width           =   540
                  _ExtentX        =   953
                  _ExtentY        =   635
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
                  Text            =   "###"
                  MaxLength       =   3
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
               Begin Cis3D_v60.CIS3D CIS3D26 
                  Height          =   255
                  Left            =   870
                  Top             =   120
                  Width           =   210
                  _ExtentX        =   370
                  _ExtentY        =   450
                  Caption         =   "～"
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cBoderStyle     =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D27 
               Height          =   510
               Left            =   15
               Top             =   15
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   900
               BackColor       =   12648384
               Caption         =   "2"
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
               cAlingnment     =   7
            End
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   345
            Index           =   3
            Left            =   6195
            Top             =   45
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   609
            ForeColor       =   0
            Caption         =   "発行済枚数"
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
      End
   End
   Begin Cis3D_v60.CIS3D B2_Area1 
      Height          =   3870
      Left            =   1215
      Top             =   6240
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   6826
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
      Begin Cis3D_v60.CIS3D CIS3D35 
         Height          =   3750
         Left            =   75
         Top             =   60
         Width           =   12150
         _ExtentX        =   21431
         _ExtentY        =   6615
         ForeColor       =   0
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
         cAlingnment     =   6
         Begin vsFlexLib.vsFlexArray vsGrid1 
            Height          =   3540
            Left            =   105
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   105
            Width           =   11955
            _Version        =   196608
            _ExtentX        =   21087
            _ExtentY        =   6244
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
            FormatString    =   "№ |仕入先 |仕入先名　　　　　　|受入|材  質　　　　　　　　　|背番号|収容器   |収容数  |発行枚数||"
            Rows            =   11
            Cols            =   11
            BackColor       =   12648447
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   0
            ScrollBars      =   0
            RowHeightMin    =   320
            GridLines       =   2
            SelectionMode   =   1
         End
      End
   End
End
Attribute VB_Name = "CKNZ0210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   材料・発注かんばん個別発行
'**       フォームID    :   CKNZ0210
'**       処理概要      :
'**
'**       作  成  日    :   2004/02/18  By CIS SIMPR-A
'**       変  更  日    :   2005/03/01  By CIS 最終発行連番を流動管理テーブルより表示に変更
'**       変  更  日    :   2008/04/11  By CIS 背番号８桁対応
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim CisVsGrid       As New CisVsGrid3

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
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 9
        .InitGet = False
        .Init
    End With
    Call GridInit
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    FormAct = Not FormAct
' 自社情報取得
    Kn_Jisya = ""
    Kn_JisyaNM = ""
    Kn_JisyaKJNM = ""
    
    Kn_Jisya = RTrim(gCompany)
    
' 自社伝票発行用(工場)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 7
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 7
        IRN.oyian = "0"
        IRN.oyian2 = "00"
    End If
    ' 工場
    Kn_Jisya = Kn_Jisya & RTrim(IRN.oyian)
    
    gSL_Select = "Select * From 取引先マスタ "
    gSL_Select = gSL_Select & " Where 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & Kn_Jisya & "'"
    If TRMRead(gSL_Select, 1) Then
       Kn_JisyaNM = TRM.略称
       Kn_JisyaKJNM = TRM.工場名
    End If

'
    Call GridSet
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
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    Select Case Key_Code
           Case vbKeyF1:       If PB_Add.Visible Then Call PB_Add_Click    '【新規】
           Case vbKeyF3:       If PB_Del.Visible Then Call PB_Del_Click    '【削除】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyF8:       If PB_Prt.Visible Then Call PB_Prt_Click    '【印刷】
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
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
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    Cancel = 1
    Call PB_END_Click
End Sub

'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
'+---------------------+
'+   ﾃﾞｰﾀﾍﾞｰｽ切断
'+---------------------+
    Call CisDB.DBDISConnect
'+---------------------+
'+   ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
'+---------------------+
    Call CisFrm.UnLoadDisp
    End
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
   Call ReturnPress
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
   Call CisFrm.MousePT(11)
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      iH1_Seban.SetFocus
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    追　加  キ ー(F1)        +
'+----------------------------+
Private Sub PB_Add_Click()
    If SyoriKB <> "ADD" Then
       K_Sykbnm = "追加"
       K_Sykbnm.ForeColor = PB_Add.BtnForeColor
       SyoriKB = "ADD"
       ProcHB = "H1"
       vsGrid1.HighLight = flexHighlightNever
       Call DispChange(ProcHB)
    End If
End Sub
'+----------------------------+
'+    削　除  キ ー(F3)        +
'+----------------------------+
Private Sub PB_Del_Click()
    If SyoriKB <> "DEL" Then
       K_Sykbnm = "削除"
       K_Sykbnm.ForeColor = PB_Del.BtnForeColor
       SyoriKB = "DEL"
       ProcHB = "B2"
       Call DispChange(ProcHB)
       vsGrid1.HighLight = flexHighlightAlways
       vsGrid1.SetFocus
    End If
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    
    Dim wk_Seban        As String * 5
    Dim wk_Torcd        As String * 7
    Dim wk_Ukeir        As String * 2
    Dim wk_Zai          As String * 25
    Dim wk_Ita          As String * 7
    Dim wk_Haba         As String * 7
    Dim wk_Nagasa       As String * 7
    
    wk_Seban = iH1_Seban
    wk_Torcd = iH1_Torcd
    wk_Ukeir = iH1_Ukeir
    wk_Zai = iH1_Zai
    wk_Ita = Format(CisFun.Val2(iH1_Ita), "#0.00")
    wk_Haba = Format(CisFun.Val2(iH1_Haba), "#0.00")
    wk_Nagasa = Format(CisFun.Val2(iH1_Nagasa), "#0.00")
    
    gSL_Select = "Select ZM.材料名,RTRIM(SY.値名称) 手配区分名,RTrim(TR.略称) メーカー名,"
    gSL_Select = gSL_Select & " ZT.*"
    gSL_Select = gSL_Select & "  From 材料取引先マスタ ZT"
    gSL_Select = gSL_Select & "   Left Outer Join 材料マスタ ZM "
    gSL_Select = gSL_Select & "   On ZM.材料管理番号 = ZT.材料管理番号"
    gSL_Select = gSL_Select & "  Left Outer Join 名称マスタ SY"
    gSL_Select = gSL_Select & "   ON  SY.区分名称 = '手配区分'"
    gSL_Select = gSL_Select & "   And SY.区分タイプ = ''"
    gSL_Select = gSL_Select & "   And SY.値 = ZT.手配区分"
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON  TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "   And TR.取引先CD = ZT.ﾒｰｶｰ"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
    'gSL_Select = gSL_Select & " Where ZT.背番号 + REPLICATE(' ',5 - LEN(ZT.背番号)) + "
    gSL_Select = gSL_Select & " Where ZT.背番号 + REPLICATE(' ',8 - LEN(ZT.背番号)) + "
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
    gSL_Select = gSL_Select & "       ZT.取引先 + REPLICATE(' ',7 - LEN(ZT.取引先)) + "
    gSL_Select = gSL_Select & "       ZT.受入 + REPLICATE(' ',2 - LEN(ZT.受入)) + "
    gSL_Select = gSL_Select & "       ZT.材質 + REPLICATE(' ',25 - LEN(ZT.材質)) + "
    gSL_Select = gSL_Select & "       Cast(ZT.板厚 as varchar(7)) + REPLICATE(' ',7 - LEN(Cast(ZT.板厚 as varchar(7)))) + "
    gSL_Select = gSL_Select & "       Cast(ZT.幅 as varchar(7)) + REPLICATE(' ',7 - LEN(Cast(ZT.幅 as varchar(7)))) + "
    gSL_Select = gSL_Select & "       Cast(ZT.長さ as varchar(7)) + REPLICATE(' ',7 - LEN(Cast(ZT.長さ as varchar(7))))  "
    gSL_Select = gSL_Select & " < '"
    gSL_Select = gSL_Select & wk_Seban & wk_Torcd & wk_Ukeir & wk_Zai & wk_Ita & wk_Haba & wk_Nagasa
    gSL_Select = gSL_Select & "'"
    gSL_Select = gSL_Select & " And SK区分 = 0 And 手配区分 <> '2'"
    gSL_Select = gSL_Select & " ORDER BY ZT.背番号 + ZT.取引先 + ZT.受入 + ZT.材質 + "
    gSL_Select = gSL_Select & "       Cast(ZT.板厚 as varchar(7)) + Cast(ZT.幅 as varchar(7)) +Cast(ZT.長さ as varchar(7)) Desc "

    Call DBInput("Back")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Seban.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    Dim wk_Seban        As String * 5
    Dim wk_Torcd        As String * 7
    Dim wk_Ukeir        As String * 2
    Dim wk_Zai          As String * 25
    Dim wk_Ita          As String * 7
    Dim wk_Haba         As String * 7
    Dim wk_Nagasa       As String * 7
    
    wk_Seban = iH1_Seban
    wk_Torcd = iH1_Torcd
    wk_Ukeir = iH1_Ukeir
    wk_Zai = iH1_Zai
    wk_Ita = Format(CisFun.Val2(iH1_Ita), "#0.00")
    wk_Haba = Format(CisFun.Val2(iH1_Haba), "#0.00")
    wk_Nagasa = Format(CisFun.Val2(iH1_Nagasa), "#0.00")
    
    gSL_Select = "Select ZM.材料名,RTRIM(SY.値名称) 手配区分名,RTrim(TR.略称) メーカー名,"
    gSL_Select = gSL_Select & " ZT.*"
    gSL_Select = gSL_Select & "  From 材料取引先マスタ ZT"
    gSL_Select = gSL_Select & "   Left Outer Join 材料マスタ ZM "
    gSL_Select = gSL_Select & "   On ZM.材料管理番号 = ZT.材料管理番号"
    gSL_Select = gSL_Select & "  Left Outer Join 名称マスタ SY"
    gSL_Select = gSL_Select & "   ON  SY.区分名称 = '手配区分'"
    gSL_Select = gSL_Select & "   And SY.区分タイプ = ''"
    gSL_Select = gSL_Select & "   And SY.値 = ZT.手配区分"
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON  TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "   And TR.取引先CD = ZT.ﾒｰｶｰ"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
    'gSL_Select = gSL_Select & " Where ZT.背番号 + REPLICATE(' ',5 - LEN(ZT.背番号)) + "
    gSL_Select = gSL_Select & " Where ZT.背番号 + REPLICATE(' ',8 - LEN(ZT.背番号)) + "
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
    gSL_Select = gSL_Select & "       ZT.取引先 + REPLICATE(' ',7 - LEN(ZT.取引先)) + "
    gSL_Select = gSL_Select & "       ZT.受入 + REPLICATE(' ',2 - LEN(ZT.受入)) + "
    gSL_Select = gSL_Select & "       ZT.材質 + REPLICATE(' ',25 - LEN(ZT.材質)) + "
    gSL_Select = gSL_Select & "       Cast(ZT.板厚 as varchar(7)) + REPLICATE(' ',7 - LEN(Cast(ZT.板厚 as varchar(7)))) + "
    gSL_Select = gSL_Select & "       Cast(ZT.幅 as varchar(7)) + REPLICATE(' ',7 - LEN(Cast(ZT.幅 as varchar(7)))) + "
    gSL_Select = gSL_Select & "       Cast(ZT.長さ as varchar(7)) + REPLICATE(' ',7 - LEN(Cast(ZT.長さ as varchar(7))))  "
    gSL_Select = gSL_Select & " > '"
    gSL_Select = gSL_Select & wk_Seban & wk_Torcd & wk_Ukeir & wk_Zai & wk_Ita & wk_Haba & wk_Nagasa
    gSL_Select = gSL_Select & "'"
    gSL_Select = gSL_Select & " And SK区分 = 0 And 手配区分 <> '2'"
    gSL_Select = gSL_Select & " ORDER BY ZT.背番号 + ZT.取引先 + ZT.受入 + ZT.材質 + "
    gSL_Select = gSL_Select & "       Cast(ZT.板厚 as varchar(7)) + Cast(ZT.幅 as varchar(7)) +Cast(ZT.長さ as varchar(7)) "

    Call DBInput("Next")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Seban.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    印　刷  キ ー(F8)       +
'+----------------------------+
Private Sub PB_Prt_Click()
    kn_Call = "C"
    kn_PrtKB = 2
    
    CKNZ0220.Show vbModal
    Unload CKNZ0220
    Set CKNZ0220 = Nothing

    Call GridSet
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Seban" Then Call Look_ZaiTori: Exit Sub
    If PB_Look.Tag = "iH1_Zai" Then Call Look_Zairyo: Exit Sub
    If PB_Look.Tag = "iH1_Ita" Then Call Look_Zairyo: Exit Sub
    If PB_Look.Tag = "iH1_Haba" Then Call Look_Zairyo: Exit Sub
    If PB_Look.Tag = "iH1_Nagasa" Then Call Look_Zairyo: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_UTorcd: Exit Sub
End Sub
'+----------------------------+
'+       仕入先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 9
    RV_Call = ""
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'+-----------------------------+
'+       取引先受入検索         +
'+-----------------------------+
Private Sub Look_UTorcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 9
    RV_Torcd = iH1_Torcd
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RV_Ukeir
    End If
End Sub
'+-----------------------------+
'+       材料検索              +
'+-----------------------------+
Private Sub Look_Zairyo()
    RV_Left = 0
    RV_Top = 0
        
    RV_Call = ""
    
    CKK0070.Show vbModal
    Unload CKK0070
    Set CKK0070 = Nothing

    If RV_Rtn Then
        iH1_Zai = RV_Zaist
        iH1_Ita = ""
        iH1_Haba = ""
        iH1_Nagasa = ""
        If RV_Size = 0 Then
           If RV_Itatu <> 0 Then iH1_Ita = CisFun.RSetFld(RV_Itatu, 0, iH1_Ita.cFormat)
           If RV_Width <> 0 Then iH1_Haba = CisFun.RSetFld(RV_Width, 0, iH1_Haba.cFormat)
           If RV_Long <> 0 Then iH1_Nagasa = CisFun.RSetFld(RV_Long, 0, iH1_Nagasa.cFormat)
        End If
    End If
End Sub
'+-----------------------------+
'+       材料取引先検索        +
'+-----------------------------+
Private Sub Look_ZaiTori()
    RV_Left = 0
    RV_Top = 0
        
    RV_Call = ""
    RVI_SK = 0
    
    CKK0077.Show vbModal
    Unload CKK0077
    Set CKK0077 = Nothing

    If RV_Rtn Then
        iH1_Seban = RV_Seban
        iH1_Torcd = RV_TorcdK
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RV_Ukeir
        iH1_Zai = RV_Zaist
        iH1_Ita = ""
        iH1_Haba = ""
        iH1_Nagasa = ""
        If RV_Size = 0 Then
           If RV_Itatu <> 0 Then iH1_Ita = CisFun.RSetFld(RV_Itatu, 0, iH1_Ita.cFormat)
           If RV_Width <> 0 Then iH1_Haba = CisFun.RSetFld(RV_Width, 0, iH1_Haba.cFormat)
           If RV_Long <> 0 Then iH1_Nagasa = CisFun.RSetFld(RV_Long, 0, iH1_Nagasa.cFormat)
        End If
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
    ' 背番号より他情報表示
    If Trim(iH1_Seban) <> "" Then
        gSL_Select = "select ZT.材質,ZT.板厚,ZT.寸法区分,ZT.板厚,ZT.幅,ZT.長さ,"
        gSL_Select = gSL_Select & "ZT.背番号,ZT.取引先,ZT.受入,ZT.最終発行連番,"
        gSL_Select = gSL_Select & "Rtrim(TR.略称) 手配先名"
        gSL_Select = gSL_Select & " From 材料取引先マスタ ZT "
        gSL_Select = gSL_Select & " Left Outer Join 取引先マスタ TR"
        gSL_Select = gSL_Select & "   On  TR.取引先区分 = 1"
        gSL_Select = gSL_Select & "   AND TR.取引先CD = ZT.取引先"
        gSL_Select = gSL_Select & " Where ZT.背番号 = '" & RTrim(iH1_Seban) & "'"
        gSL_Select = gSL_Select & "   And ZT.SK区分 = 0"
        gSL_Select = gSL_Select & "   And ZT.手配区分 <> '2'"
        If ZTMRead(gSL_Select, 0, 1) Then
           If CisDB.RecordCount = 1 Then
              iH1_Seban = RTrim(ZTM.背番号)
              iH1_Torcd = RTrim(ZTM.取引先)
              iH1_Ukeir = RTrim(ZTM.受入)
              iH1_Zai = RTrim(ZTM.材質)
              iH1_Ita = ""
              iH1_Haba = ""
              iH1_Nagasa = ""
              If ZTM.寸法区分 = 0 Then
                 If ZTM.板厚 <> 0 Then iH1_Ita = CisFun.RSetFld(ZTM.板厚, 0, iH1_Ita.cFormat)
                 If ZTM.幅 <> 0 Then iH1_Haba = CisFun.RSetFld(ZTM.幅, 0, iH1_Haba.cFormat)
                 If ZTM.長さ <> 0 Then iH1_Nagasa = CisFun.RSetFld(ZTM.長さ, 0, iH1_Nagasa.cFormat)
              End If
              Call iH1_Torcd_LostFocus
           End If
        End If
        Call ZTMClose(1)
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
    If TorNmGet(iH1_Torcd, 9) Then
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
'+  iH1_Zai
'+-----------------------------+
Private Sub iH1_Zai_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Zai_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_Ita
'+-----------------------------+
Private Sub iH1_Ita_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Ita_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_Haba
'+-----------------------------+
Private Sub iH1_Haba_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Haba_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_Nagasa
'+-----------------------------+
Private Sub iH1_Nagasa_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Nagasa_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub

'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
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
    H1Mode = False
    H1Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
    B2Mode = False
    B2Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
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
       Case "B2"
          B2Mode = True
          B2Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B2_Area1.Enabled = B2Mode
    B2_Area1.BackColor = B2Color
    
    PB_Look.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = B1Mode
    PB_Del.Visible = H1Mode
    PB_Add.Visible = B2Mode
    
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
'        CisVsGrid.Clear
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        Call DispChange(ProcHB)
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
    If ProcHB = "B2" Then
        If Not DBDelete Then
           vsGrid1.SetFocus
           GoTo ReturnPress_Ed
        End If
        vsGrid1.SetFocus
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
    If Trim(iH1_Seban) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    背番号を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Seban.SetFocus
       Exit Function
    End If
    
    gSL_Select = "select ZT.* From 材料取引先マスタ ZT "
    gSL_Select = gSL_Select & " Where ZT.背番号 = '" & RTrim(iH1_Seban) & "'"
    gSL_Select = gSL_Select & "   And ZT.取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   And ZT.受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   And ZT.材質 = '" & RTrim(iH1_Zai) & "'"
    gSL_Select = gSL_Select & "   And ZT.板厚 = " & CisFun.Val2(iH1_Ita)
    gSL_Select = gSL_Select & "   And ZT.幅 = " & CisFun.Val2(iH1_Haba)
    gSL_Select = gSL_Select & "   And ZT.長さ = " & CisFun.Val2(iH1_Nagasa)
    If Not ZTMRead(gSL_Select, 1, 1) Then
        gSL_Select = "select ZT.材質,ZT.板厚,ZT.寸法区分,ZT.板厚,ZT.幅,ZT.長さ,"
        gSL_Select = gSL_Select & "ZT.背番号,ZT.取引先,ZT.受入,ZT.最終発行連番,"
        gSL_Select = gSL_Select & "Rtrim(TR.略称) 手配先名"
        gSL_Select = gSL_Select & " From 材料取引先マスタ ZT "
        gSL_Select = gSL_Select & " Left Outer Join 取引先マスタ TR"
        gSL_Select = gSL_Select & "   On  TR.取引先区分 = 1"
        gSL_Select = gSL_Select & "   AND TR.取引先CD = ZT.取引先"
        gSL_Select = gSL_Select & " Where ZT.背番号 = '" & RTrim(iH1_Seban) & "'"
        gSL_Select = gSL_Select & "   And ZT.SK区分 = 0"
        gSL_Select = gSL_Select & "   And ZT.手配区分 <> '2'"
        If ZTMRead(gSL_Select, 0, 1) Then
           If CisDB.RecordCount > 1 Then
              CKNZ0215.Show vbModal
              Unload CKNZ0215
              Set CKNZ0215 = Nothing
              If RV_Rtn Then
                 iH1_Torcd = RTrim(RV_TorcdK)
                 iH1_Ukeir = RTrim(RV_Ukeir)
                 iH1_Zai = RTrim(RV_Zaist)
                 iH1_Ita = ""
                 iH1_Haba = ""
                 iH1_Nagasa = ""
                 If RV_Size = 0 Then
                    If RV_Itatu <> 0 Then iH1_Ita = CisFun.RSetFld(RV_Itatu, 0, iH1_Ita.cFormat)
                    If RV_Width <> 0 Then iH1_Haba = CisFun.RSetFld(RV_Width, 0, iH1_Haba.cFormat)
                    If RV_Long <> 0 Then iH1_Nagasa = CisFun.RSetFld(RV_Long, 0, iH1_Nagasa.cFormat)
                 End If
                 Call iH1_Torcd_LostFocus
              Else
                 iH1_Seban.SetFocus
                 Exit Function
              End If
           End If
        End If
        Call ZTMClose(1)
    End If
    
    If Trim(iH1_Torcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    手配先を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    End If
    If Not TorNmGet(iH1_Torcd, 9) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    手配先　取引先マスタ未登録            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    End If
    If Trim(iH1_Ukeir) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    受入を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Ukeir.SetFocus
       Exit Function
    End If
    gSL_Select = "select * From 取引先受入マスタ "
    gSL_Select = gSL_Select & " WHERE 取引先CD = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   And 取引先区分 = 1"
    gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(iH1_Ukeir) & "'"
    If Not TUMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    受入未登録            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Ukeir.SetFocus
       Exit Function
    End If
    If Trim(iH1_Zai) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    材質を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Zai.SetFocus
       Exit Function
    End If

' マスタ入力 SELECT文 ｾｯﾄ
    gSL_Select = "Select ZM.材料名,RTRIM(SY.値名称) 手配区分名,RTrim(TR.略称) メーカー名,"
    gSL_Select = gSL_Select & " ZT.*"
    gSL_Select = gSL_Select & "  From 材料取引先マスタ ZT"
    gSL_Select = gSL_Select & "   Left Outer Join 材料マスタ ZM "
    gSL_Select = gSL_Select & "   On ZM.材料管理番号 = ZT.材料管理番号"
    gSL_Select = gSL_Select & "  Left Outer Join 名称マスタ SY"
    gSL_Select = gSL_Select & "   ON  SY.区分名称 = '手配区分'"
    gSL_Select = gSL_Select & "   And SY.区分タイプ = ''"
    gSL_Select = gSL_Select & "   And SY.値 = ZT.手配区分"
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON  TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "   And TR.取引先CD = ZT.ﾒｰｶｰ"
    gSL_Select = gSL_Select & " Where ZT.背番号 = '" & RTrim(iH1_Seban) & "'"
    gSL_Select = gSL_Select & "   And ZT.取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   And ZT.受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   And ZT.材質 = '" & RTrim(iH1_Zai) & "'"
    gSL_Select = gSL_Select & "   And ZT.板厚 = " & CisFun.Val2(iH1_Ita)
    gSL_Select = gSL_Select & "   And ZT.幅 = " & CisFun.Val2(iH1_Haba)
    gSL_Select = gSL_Select & "   And ZT.長さ = " & CisFun.Val2(iH1_Nagasa)
    
    If Not DBInput("INV") Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    材料取引先マスタ未登録            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_Seban.SetFocus
        Exit Function
    Else
        If ZTM.SK区分 <> 0 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     かんばんを指定して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Seban.SetFocus
            Exit Function
        End If
        If ZTM.手配区分 = "2" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     手配先は'社内'以外を指定して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Seban.SetFocus
            Exit Function
        End If
    End If
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Dim NoStr       As String
    Body1Chk = False
    
'    ' 追加枚数のみ（１）
    If Not Back_HSu(1).Visible And Not Back_SSu(0).Visible And Not Back_SSu(1).Visible Then
       If CisFun.Val2(iB1_HSu(0)) = 0 Then
          With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     追加枚数を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
          End With
          Exit Function
       End If
    End If
    If Not Back_HSu(1).Visible And Back_SSu(0).Visible And Not Back_SSu(1).Visible Then
       If CisFun.Val2(iB1_HSu(0)) = 0 And _
          CisFun.Val2(iB1_SSSu(0)) = 0 And _
          CisFun.Val2(iB1_ESSu(0)) = 0 Then
          With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     追加枚数又は再発行№を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
          End With
          Exit Function
       End If
    End If
    If Trim(B1lb_SHYmd) <> "" Then
       If CisFun.Val2(iB1_HSu(0)) = 0 And _
          CisFun.Val2(iB1_SSSu(0)) = 0 And _
          CisFun.Val2(iB1_ESSu(0)) = 0 And _
          CisFun.Val2(iB1_HSu(1)) = 0 And _
          CisFun.Val2(iB1_SSSu(1)) = 0 And _
          CisFun.Val2(iB1_ESSu(1)) = 0 Then
          With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     追加枚数又は再発行№を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
          End With
          Exit Function
       End If
     End If
    For gInt = 0 To 1
        If gInt = 0 Then
           NoStr = "①"
        Else
           NoStr = "②"
        End If
        If Trim(B1lb_SHYmd) = "" Then
           NoStr = ""
        End If
        '+------------------------------------------- 2005/03/01 Insert
        If Back_SSu(gInt).Visible And _
           CisFun.Val2(iB1_HSu(gInt)) <> 0 And _
           (CisFun.Val2(iB1_SSSu(gInt)) <> 0 Or _
           CisFun.Val2(iB1_ESSu(gInt)) <> 0) Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "     追加枚数 又は 再発行№" & NoStr & "のどちらか片方にて入力して下さい          "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB1_SSSu(gInt).SetFocus
           Exit Function
        End If
        '+------------------------------------------- 2005/03/01 Insert End
        If Back_SSu(gInt).Visible And _
           CisFun.Val2(iB1_HSu(gInt)) = 0 And _
           CisFun.Val2(iB1_SSSu(gInt)) = 0 And _
           CisFun.Val2(iB1_ESSu(gInt)) <> 0 Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "     再発行№" & NoStr & "(開始)を入力して下さい            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB1_SSSu(gInt).SetFocus
           Exit Function
        End If
        If Back_SSu(gInt).Visible And _
           CisFun.Val2(iB1_HSu(gInt)) = 0 And _
           CisFun.Val2(iB1_ESSu(gInt)) = 0 And _
           CisFun.Val2(iB1_SSSu(gInt)) <> 0 Then
           With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "     再発行№" & NoStr & "(終了)を入力して下さい            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
            End With
            iB1_ESSu(gInt).SetFocus
            Exit Function
         End If
         If CisFun.Val2(iB1_SSSu(gInt)) > CisFun.Val2(iB1_ESSu(gInt)) Then
            With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "     再発行№" & NoStr & "を正しく入力して下さい（開始＞終了）            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
            End With
            iB1_ESSu(gInt).SetFocus
            Exit Function
         End If
         If CisFun.Val2(iB1_ESSu(gInt)) > CisFun.Val2(B1lb_EndSu) Then
            With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "     再発行№" & NoStr & "は最終№以下を指定して下さい            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
            End With
            iB1_ESSu(gInt).SetFocus
            Exit Function
         End If
        ' 再発行№チェック(かんばん材料管理テーブル)
        If CisFun.Val2(iB1_SSSu(gInt)) <> 0 Then
           For gLong = CisFun.Val2(iB1_SSSu(gInt)) To CisFun.Val2(iB1_ESSu(gInt))
               gSL_Select = "Select * From かんばん材料管理テーブル"
               gSL_Select = gSL_Select & " Where 材料管理番号 = " & CisFun.Val2(iH1_Seban.Tag)
               gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(iH1_Torcd) & "'"
               gSL_Select = gSL_Select & " And   受入 = '" & RTrim(iH1_Ukeir) & "'"
               gSL_Select = gSL_Select & " And   発行NO = " & gLong & ""
               If KZKRead(gSL_Select, 1) Then
                  If gInt + 1 <> KZK.収容区分 Then
                    With CisFun
                      .MB_Lines = 4
                      .MB_MSG(2) = "     再発行№(" & StrConv(Format(gLong, "000"), vbWide) & ")の収容数が前回発行と異なります            "
                      .MB_Title = ""
                      .MB_Button = Error
                      .MBOX
                    End With
                    iB1_SSSu(gInt).SetFocus
                    Exit Function
                  End If
               End If
           Next gLong
        End If
    Next gInt
    
    Body1Chk = True
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    Call HeadBodyClear("B")
    
    If Not ZTMRead(gSL_Select, 1) Then
        If InputKey = "Back" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     先頭レコードです            "
                .MB_Title = "【後退】"
                .MB_Button = OK
                .MBOX
            End With
            GoTo DBInput_Ed
        End If
        If InputKey = "Next" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     最終レコードです            "
                .MB_Title = "【前進】"
                .MB_Button = OK
                .MBOX
            End With
            GoTo DBInput_Ed
        End If
    Else
        iH1_Seban = RTrim(ZTM.背番号)
        iH1_Seban.Tag = ZTM.材料管理番号
        iH1_Torcd = RTrim(ZTM.取引先)
        iH1_Ukeir = RTrim(ZTM.受入)
        iH1_Zai = RTrim(ZTM.材質)
        iH1_Ita = ""
        iH1_Haba = ""
        iH1_Nagasa = ""
        If ZTM.寸法区分 = 0 Then
           If ZTM.板厚 <> 0 Then iH1_Ita = CisFun.RSetFld(ZTM.板厚, 0, iH1_Ita.cFormat)
           If ZTM.幅 <> 0 Then iH1_Haba = CisFun.RSetFld(ZTM.幅, 0, iH1_Haba.cFormat)
           If ZTM.長さ <> 0 Then iH1_Nagasa = CisFun.RSetFld(ZTM.長さ, 0, iH1_Nagasa.cFormat)
        End If
        Call BodySet("")
        DBInput = True
    End If
DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function

'**************************************
'*          BODY 内容ｾｯﾄ処理           *
'**************************************
Private Sub BodySet(SetType As String)
    If ZTM.SK区分 <> 0 Then Exit Sub

    With ZTM
        If Trim(.サイクル) <> "" Then
            B1lb_Cycle = CisFun.RSetFld(Mid(.サイクル, 1, 1), 1, "#") & "-" & _
                         CisFun.RSetFld(Mid(.サイクル, 2, 2), 2, "#") & "-" & _
                         CisFun.RSetFld(Mid(.サイクル, 4, 2), 2, "#")
        End If
        B1lb_Zainm = .材料名
        If .手配区分 = "1" Then
           Back_Torcd = "納入先（直送）"
        Else
           Back_Torcd = "ﾒｰｶｰ（手配商社)"
        End If
        B1lb_Maker = .メーカー名
        B1lb_MUke = .メーカー受入
        B1lb_KTorcd = .契約先
        B1lb_TKbn = .手配区分名
        B1lb_Syuki(0) = .収容器1
        If .収容数1 <> 0 Then
           B1lb_Syuyo(0) = Format(.収容数1, "#,##0.000")
        End If
        B1lb_Syuki(1) = .収容器2
        If .収容数2 <> 0 Then
           B1lb_Syuyo(1) = Format(.収容数2, "#,##0.000")
        End If
        
        If Trim(.変更日) <> "" Then
           B1lb_SHYmd = Mid(.変更日, 1, 4) & "/" & Mid(.変更日, 5, 2) & "/" & _
                        Mid(.変更日, 7, 2)
           B1lb_SHBin = .変更便
           Back_HSu(1).Visible = True
           Back_SSu(1).Visible = True
        Else
           Back_HSu(1).Visible = False
           Back_SSu(1).Visible = False
        End If
        If .最終発行連番 <> 0 Then
           B1lb_EndSu = Format(.最終発行連番, "000")
           Back_SSu(0).Visible = True
        Else
           Back_SSu(0).Visible = False
           Back_SSu(1).Visible = False
        End If
    End With
    Call iH1_Torcd_LostFocus
    
    '------------------------------------------------------------------- 2005/03/01 Insert
    gSL_Select = "Select Max(発行NO) 発行NO From かんばん材料管理テーブル "
    gSL_Select = gSL_Select & " Where 材料管理番号 = " & ZTM.材料管理番号
    gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(ZTM.取引先) & "'"
    gSL_Select = gSL_Select & " And   受入 = '" & RTrim(ZTM.受入) & "'"
    If KZKRead(gSL_Select, 1) Then
        If KZK.発行NO > CisFun.Val2(B1lb_EndSu) Then
           B1lb_EndSu = Format(KZK.発行NO, "000")
           Back_SSu(0).Visible = True
        End If
    End If
    '------------------------------------------------------------------- 2005/03/01 Insert End
    gSL_Select = "Select Count(*) 件数 From かんばん材料管理テーブル "
    gSL_Select = gSL_Select & " Where 材料管理番号 = " & ZTM.材料管理番号
    gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(ZTM.取引先) & "'"
    gSL_Select = gSL_Select & " And   受入 = '" & RTrim(ZTM.受入) & "'"
    gSL_Select = gSL_Select & " And   収容区分 = 1"
    If KZKRead(gSL_Select, 1) Then
        B1lb_HZSu(0) = Format(KZK.件数, "###")
    End If
    
    If Back_HSu(1).Visible Then
        gSL_Select = "Select Count(*) 件数 From かんばん材料管理テーブル "
        gSL_Select = gSL_Select & " Where 材料管理番号 = " & ZTM.材料管理番号
        gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(ZTM.取引先) & "'"
        gSL_Select = gSL_Select & " And   受入 = '" & RTrim(ZTM.受入) & "'"
        gSL_Select = gSL_Select & " And   収容区分 = 2"
        If KZKRead(gSL_Select, 1) Then
            B1lb_HZSu(1) = Format(KZK.件数, "###")
        End If
    End If
    
End Sub
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+    削除 処理          +
'+----------------------+
Private Function DBDelete() As Boolean
    Dim wHMai As Long
    DBDelete = False
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     選択行を削除します      "
        .MB_Title = "削除処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
'
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

    With CisDB
        .SQL = "Delete From かんばん発行テーブル"
        .SQL = .SQL & " Where かんばん種類 = 2"
        .SQL = .SQL & " And   背番号 = '" & RTrim(vsGrid1.TextMatrix(vsGrid1.Row, 5)) & "'"
        .SQL = .SQL & " And   材料管理番号 = " & CisFun.Val2(vsGrid1.TextMatrix(vsGrid1.Row, 9))
        .SQL = .SQL & " And   取引先 = '" & RTrim(vsGrid1.TextMatrix(vsGrid1.Row, 1)) & "'"
        .SQL = .SQL & " And   受入 = '" & RTrim(vsGrid1.TextMatrix(vsGrid1.Row, 3)) & "'"
        .SQL = .SQL & " And   収容数 = " & CisFun.Val2(vsGrid1.TextMatrix(vsGrid1.Row, 7))
        If Not .DBExec Then GoTo DBDelete_Err
    End With
    With vsGrid1
    ' 最終発行連番を求める
        wHMai = 0
        gSL_Select = "Select Max(発行NO) as 発行NO From かんばん材料管理テーブル"
        gSL_Select = gSL_Select & " Where 材料管理番号 = " & CisFun.Val2(.TextMatrix(.Row, 9))
        gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(.TextMatrix(.Row, 1)) & "'"
        gSL_Select = gSL_Select & " And   受入 = '" & RTrim(.TextMatrix(.Row, 3)) & "'"
        If KZKRead(gSL_Select, 1) Then
           wHMai = KZK.発行NO
        End If
    ' 最終発行連番を求める
        gSL_Select = "Select Max(発行NO) as 発行NO From かんばん発行テーブル"
        gSL_Select = gSL_Select & " Where かんばん種類 = 2"
        gSL_Select = gSL_Select & " And   材料管理番号 = " & CisFun.Val2(.TextMatrix(.Row, 9))
        gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(.TextMatrix(.Row, 1)) & "'"
        gSL_Select = gSL_Select & " And   受入 = '" & RTrim(.TextMatrix(.Row, 3)) & "'"
        gSL_Select = gSL_Select & " And   発行区分 = 0"
    End With
    If KHTRead(gSL_Select, 1) Then
       If KHT.発行NO > wHMai Then wHMai = KHT.発行NO
       With CisDB
            .SQL = "Update 材料取引先マスタ Set  "
            .SQL = .SQL & "最終発行連番 = " & wHMai & ","
            .SQL = .SQL & "更新日 = GETDATE(),"
            .SQL = .SQL & "更新者 = '" & gTanto & "',"
            .SQL = .SQL & "更新端末 = SubString(Host_Name(),1,20)"
            .SQL = .SQL & " Where 材料管理番号 = " & CisFun.Val2(vsGrid1.TextMatrix(vsGrid1.Row, 9))
            .SQL = .SQL & "   And 取引先 = '" & RTrim(vsGrid1.TextMatrix(vsGrid1.Row, 1)) & "'"
            .SQL = .SQL & "   And 受入 = '" & RTrim(vsGrid1.TextMatrix(vsGrid1.Row, 3)) & "'"
            If Not .DBExec Then GoTo DBDelete_Err
       End With
    End If
'
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
'
    Call GridSet
'
    DBDelete = True
    Exit Function
DBDelete_Err:
    CisDB.DBTran (TransRollback) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加 処理          +
'+----------------------+
Private Function DBPut() As Boolean
    DBPut = False
    Dim AddMai  As Long
    Dim PrtStart    As Long
    Dim PrtEnd      As Long
    Dim wTorNm      As String
    Dim wwTorcd     As String
    Dim wwNTorcd    As String
    Dim wNTorNm     As String
    Dim PrtKb       As Integer
    Dim wPutFlg     As Boolean
    Dim wNTime(8)   As String
    Dim wTaniNM     As String
    Dim wTTaniNm    As String
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     発行データ作成処理を行います      "
        .MB_Title = "処理確認"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err
'   材料取引先情報
    gSL_Select = "select ZT.* From 材料取引先マスタ ZT "
    gSL_Select = gSL_Select & " Where ZT.材料管理番号 = " & CisFun.Val2(iH1_Seban.Tag)
    gSL_Select = gSL_Select & "   And ZT.取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   And ZT.受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   And ZT.背番号 = '" & RTrim(iH1_Seban) & "'"
    gSL_Select = gSL_Select & "   And ZT.SK区分 = 0"
    If Not ZTMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     材料取引先マスタ　異常      "
            .MB_Title = ""
            .MB_Button = Error
            Call .MBOX
        End With
        Exit Function
    End If
    ' 品番情報
    gSL_Select = "Select * From 材料マスタ "
    gSL_Select = gSL_Select & " Where 材料管理番号 = " & ZTM.材料管理番号
    If Not ZRMRead(gSL_Select, 1) Then
       Call ItemsClearZRM
    End If
    ' 使用品番情報
    gSL_Select = "Select 使用品番 From 材料品番マスタ"
    gSL_Select = gSL_Select & " Where 材料管理番号 = " & ZHM.材料管理番号 & ""
    gSL_Select = gSL_Select & " Order By 使用品番"
    If Not ZHMRead(gSL_Select, 1) Then
       Call ItemsClearZHM
    End If
    ' 納入先情報（メーカー)
    wNTorNm = "": wwNTorcd = ""
    If Trim(ZTM.メーカー) <> "" Then
        gSL_Select = "Select * From 取引先マスタ "
        gSL_Select = gSL_Select & " Where 取引先区分 = 1"
        gSL_Select = gSL_Select & "   AND 取引先CD = '" & ZTM.メーカー & "'"
        If Not TRMRead(gSL_Select, 1) Then
           Call ItemsClearTRM
        End If
        wNTorNm = TRM.略称
        wwNTorcd = TRM.取引先
        If Trim(TRM.工場CD) <> "" Then
           wwNTorcd = wwNTorcd & "-" & TRM.工場CD
        End If
    End If
    ' 取引先情報
    gSL_Select = "Select * From 取引先マスタ "
    gSL_Select = gSL_Select & " Where 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & ZTM.取引先 & "'"
    If Not TRMRead(gSL_Select, 1) Then
       Call ItemsClearTRM
    End If
    wTorNm = TRM.略称
    wwTorcd = TRM.取引先
    If Trim(TRM.工場CD) <> "" Then
       wwTorcd = wwTorcd & "-" & TRM.工場CD
    End If
    ' 取引先受入情報
    gSL_Select = "Select * From 取引先受入マスタ "
    gSL_Select = gSL_Select & " Where 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & ZTM.取引先 & "'"
    gSL_Select = gSL_Select & "   AND 受入 = '" & ZTM.受入 & "'"
    If Not TUMRead(gSL_Select, 1) Then
       Call ItemsClearTUM
    End If
    ' 単位名
    wTaniNM = ""
    If MeisyoGet("単位区分", ZRM.単位) Then
       wTaniNM = RTrim(SYM_Meisyo)
    End If
    wTTaniNm = ""
    If MeisyoGet("単位区分", ZRM.手配単位) Then
       wTTaniNm = RTrim(SYM_Meisyo)
    End If
    ' かんばん発行情報出力
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' 発行済データ削除
    With CisDB
        .SQL = "Delete かんばん発行テーブル "
        .SQL = .SQL & "    Where かんばん種類 = 2"
        .SQL = .SQL & "      And 発行済 = 1"
        If Not .DBExec Then GoTo DBPut_Err
    End With
' ( 追加発行 )
    For gInt = 0 To 1
        PrtStart = 0: PrtEnd = 0
        AddMai = CisFun.Val2(iB1_HSu(gInt))
        If AddMai <> 0 Then
           PrtStart = CisFun.Val2(B1lb_EndSu) + 1
           PrtEnd = PrtStart + AddMai - 1
           PrtKb = 0        ' 追加発行
        End If
        If CisFun.Val2(iB1_SSSu(gInt)) <> 0 Then
           PrtStart = CisFun.Val2(iB1_SSSu(gInt))
           PrtKb = 1        ' 再発行
        End If
        If CisFun.Val2(iB1_ESSu(gInt)) <> 0 Then
           PrtEnd = CisFun.Val2(iB1_ESSu(gInt))
           PrtKb = 1        ' 再発行
        End If
        If PrtEnd = 0 Then
           PrtEnd = PrtStart
        End If
        
        If PrtStart <> 0 Then
            For gLong = PrtStart To PrtEnd
                wPutFlg = False
                gSL_Select = "Select * From かんばん発行テーブル"
                gSL_Select = gSL_Select & " Where かんばん種類 = 2"
                gSL_Select = gSL_Select & " And   材料管理番号 = " & ZTM.材料管理番号
                gSL_Select = gSL_Select & " And   背番号 = '" & RTrim(ZTM.背番号) & "'"
                gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(ZTM.取引先) & "'"
                gSL_Select = gSL_Select & " And   受入 = '" & RTrim(ZTM.受入) & "'"
                gSL_Select = gSL_Select & " And   発行NO = " & gLong
                If Not KHTRead(gSL_Select, 1) Then
                    Call ItemsClearKHT
                    With KHT
                        .かんばん種類 = 2
                        .品番 = Format(ZTM.材料管理番号, "0000000000")
                        .取引先 = ZTM.取引先
                        .受入 = ZTM.受入
                        .発行NO = gLong
                        .材料管理番号 = ZTM.材料管理番号
                    End With
                    wPutFlg = True
                End If
                With KHT
                    .発行種類 = 2
                    .背番号 = ZTM.背番号
                    .発行区分 = PrtKb
                    .表示品番 = ZRM.表示寸法
                    .品目 = 2
                    .受入名 = TUM.受入名称
                    .収容器 = B1lb_Syuki(gInt)
                    .収容数 = CisFun.Val2(B1lb_Syuyo(gInt))
                    .収容区分 = gInt + 1
                    .手配区分 = ZTM.手配区分
                    .手配区分名 = RTrim(B1lb_TKbn)
                    .アドレス = ZTM.アドレス
                    .取引先CD = wwTorcd
                    .取引先名 = TRM.略称
                    .取引先名1 = CisFun.Mid2(TRM.略称, 1, 10)
                    .取引先名2 = CisFun.Mid2(TRM.略称, 11, 10)
                    .取引先工場名 = TRM.工場名
                    .取引先工場名1 = CisFun.Mid2(TRM.工場名, 1, 10)
                    .取引先工場名2 = CisFun.Mid2(TRM.工場名, 11, 10)
                    If ZTM.手配区分 = "1" Then
                        .納入先固定 = "直送先"
                        .納入先CD = wwNTorcd
                        .納入先 = ZTM.メーカー
                        .納入先受入 = ZTM.メーカー受入
                        .納入先名 = wNTorNm
                        .納入先名1 = CisFun.Mid2(wNTorNm, 1, 10)
                        .納入先名2 = CisFun.Mid2(wNTorNm, 11, 10)
                    Else
                        If Trim(ZTM.メーカー) <> "" Then
                           .納入先固定 = "手配商社"
                           .納入先CD = wwNTorcd
                           .納入先 = ZTM.メーカー
                           .納入先受入 = ZTM.メーカー受入
                           .納入先名 = wNTorNm
                           .納入先名1 = CisFun.Mid2(wNTorNm, 1, 10)
                           .納入先名2 = CisFun.Mid2(wNTorNm, 11, 10)
                           .メーカ固定 = "手配商社"
                           .メーカCD = wwNTorcd
                           .メーカ = ZTM.メーカー
                           .メーカ受入 = ZTM.メーカー受入
                           .メーカ名 = wNTorNm
                           .メーカ名1 = CisFun.Mid2(wNTorNm, 1, 10)
                           .メーカ名2 = CisFun.Mid2(wNTorNm, 11, 10)
                        End If
                    End If
                    If Trim(ZTM.サイクル) <> "" Then
                       .サイクル = B1lb_Cycle
                    End If
                    .置場 = ZTM.置場
                    .自社CD = Kn_Jisya
                    .自社名1 = CisFun.Mid2(Kn_JisyaNM, 1, 20)
                    .自社名2 = CisFun.Mid2(Kn_JisyaKJNM, 1, 14)
                    .作成者 = gTanto
                    If CisFun.Len2(ZRM.材質) < 19 Then
                       .材質18 = ZRM.材質
                       .材質 = ""
                    Else
                       .材質 = ZRM.材質
                       .材質18 = ""
                    End If
                    .寸法区分 = ZRM.寸法区分
                    .板厚 = ZRM.板厚
                    .幅 = ZRM.幅
                    .長さ = ZRM.長さ
                    .材料名 = ZRM.材料名
                    .使用品番 = ZHM.使用品番
                    .単位名 = CisFun.Mid2(wTaniNM, 1, 4)
                    .手配単位名 = CisFun.Mid2(wTTaniNm, 1, 4)
                    .実数区分 = ZRM.実数区分
                    .材種コード = ZRM.材種コード
                    .備考 = ZTM.備考
                
                    .納入時間1 = wNTime(1)
                    .納入時間2 = wNTime(2)
                    .納入時間3 = wNTime(3)
                    .納入時間4 = wNTime(4)
                    .納入時間5 = wNTime(5)
                    .納入時間6 = wNTime(6)
                    .納入時間7 = wNTime(7)
                    .納入時間8 = wNTime(8)
                End With
            
                If wPutFlg Then
                    Call KHTInsert
                Else
                    Call KHTUpdate
                End If
                B1lb_EndSu = Format(gLong, "000")
            Next gLong
        End If
        
        If AddMai <> 0 Then
           With CisDB
                .SQL = "Update 材料取引先マスタ Set  "
                .SQL = .SQL & "最終発行連番 = " & PrtEnd & ","
                .SQL = .SQL & "更新日 = GETDATE(),"
                .SQL = .SQL & "更新者 = '" & gTanto & "',"
                .SQL = .SQL & "更新端末 = SubString(Host_Name(),1,20)"
                .SQL = .SQL & " Where 材料管理番号 = " & ZTM.材料管理番号
                .SQL = .SQL & "   And 取引先 = '" & RTrim(ZTM.取引先) & "'"
                .SQL = .SQL & "   And 受入 = '" & RTrim(ZTM.受入) & "'"
                If Not .DBExec Then GoTo DBPut_Err
           End With
           Back_SSu(gInt).Enabled = True
        End If
    Next gInt

    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    Call GridSet
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisDB.DBTran (TransRollback) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    CisFun.ErrorBox
    End
End Function
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
        Call .FixedSet(中中, 左中, あり, 7, 1, "仕入先")
        Call .FixedSet(中中, 左中, あり, 10, 2, "仕入先名")
        Call .FixedSet(中中, 中中, あり, 2, 3, "受")
        Call .FixedSet(中中, 左中, あり, 35, 4, "材 質      板厚×幅×長さ")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
        'Call .FixedSet(中中, 左中, あり, 6, 5, "背番号")
        Call .FixedSet(中中, 左中, あり, 8, 5, "背番号")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
        Call .FixedSet(中中, 左中, あり, 10, 6, "収容器")
        Call .FixedSet(中中, 右中, あり, 11, 7, "収容数")
        Call .FixedSet(中中, 右中, あり, 4, 8, "枚数")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 9, "材料管理番号")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8")
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

    CisVsGrid.Clear

' 抽出処理 初期ｾｯﾄ
    gSL_Select = "select KT.取引先,KT.受入,KT.背番号,KT.収容器,KT.表示品番,KT.材質,"
    gSL_Select = gSL_Select & "KT.収容数,Count(KT.発行NO) 件数,KT.材料管理番号,"
    gSL_Select = gSL_Select & "    TR.略称 取引先名1"
    gSL_Select = gSL_Select & "       from かんばん発行テーブル KT"
    gSL_Select = gSL_Select & "    Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & "    On  TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "    AND TR.取引先CD = KT.取引先"
    gSL_Select = gSL_Select & "    Where KT.かんばん種類 = 2"
    gSL_Select = gSL_Select & "      And KT.発行済 = 0"
    gSL_Select = gSL_Select & "    Group by KT.取引先,KT.受入,KT.材質,KT.表示品番,KT.背番号,KT.収容器,"
    gSL_Select = gSL_Select & "             KT.材料管理番号,KT.収容数,TR.略称"
    gSL_Select = gSL_Select & "    Order by MAX(KT.作成日) Desc"
    If Not KHTRead(gSL_Select) Then GoTo GridSet_Ed
    
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    CisVsGrid.Clear
    With vsGrid1
       .Redraw = False
        If CisDB.RecordCount + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = CisDB.RecordCount + 1
        End If
    End With
    ID = 0: Ttl_Kensu = 0
    Do Until Not KHT_RDSTS
        With vsGrid1
            ID = ID + 1
            '№
            .TextMatrix(ID, 0) = Format(ID, "####")
            .TextMatrix(ID, 1) = KHT.取引先
            .TextMatrix(ID, 2) = KHT.取引先名1
            .TextMatrix(ID, 3) = KHT.受入
            .TextMatrix(ID, 4) = KHT.材質 & " " & KHT.表示品番
            .TextMatrix(ID, 5) = KHT.背番号
            .TextMatrix(ID, 6) = KHT.収容器
            .TextMatrix(ID, 7) = Format(KHT.収容数, "###0.000")
            .TextMatrix(ID, 8) = Format(KHT.件数, "###")
            
            .TextMatrix(ID, 9) = KHT.材料管理番号
            
            Ttl_Kensu = Ttl_Kensu + KHT.件数
       End With
       Call KHTReadNext
    Loop
    With vsGrid1
         iT1_Addsu = Format(Ttl_Kensu, "#,###")
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
GridSet_Ed:
    Call KHTClose
End Sub

Private Sub vsGrid1_DblClick()
    If ProcHB = "B2" Then
       If Trim(vsGrid1.TextMatrix(vsGrid1.Row, 1)) = "" Then Exit Sub
       Call ReturnPress
    End If
End Sub


