VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CKN0050 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "かんばんポケット指定"
   ClientHeight    =   11490
   ClientLeft      =   1875
   ClientTop       =   2100
   ClientWidth     =   15330
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
   ScaleHeight     =   11490
   ScaleWidth      =   15330
   Begin VB.PictureBox Picture1 
      Height          =   2685
      Left            =   285
      ScaleHeight     =   2625
      ScaleWidth      =   14625
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   8235
      Width           =   14685
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   2610
         Left            =   -15
         Top             =   15
         Width           =   14580
         _ExtentX        =   25718
         _ExtentY        =   4604
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
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   300
            Index           =   10
            Left            =   12585
            Top             =   465
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   529
            Caption         =   "エラー内容"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   300
            Index           =   11
            Left            =   12135
            Top             =   465
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   529
            Caption         =   "PNO"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   300
            Index           =   6
            Left            =   7905
            Top             =   465
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   529
            Caption         =   "エラー内容"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   300
            Index           =   7
            Left            =   7455
            Top             =   465
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   529
            Caption         =   "PNO"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   300
            Index           =   8
            Left            =   10245
            Top             =   465
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   529
            Caption         =   "エラー内容"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   300
            Index           =   9
            Left            =   9795
            Top             =   465
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   529
            Caption         =   "PNO"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   300
            Index           =   4
            Left            =   5565
            Top             =   465
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   529
            Caption         =   "エラー内容"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   300
            Index           =   5
            Left            =   5115
            Top             =   465
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   529
            Caption         =   "PNO"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   300
            Index           =   2
            Left            =   3225
            Top             =   465
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   529
            Caption         =   "エラー内容"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   300
            Index           =   3
            Left            =   2775
            Top             =   465
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   529
            Caption         =   "PNO"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   300
            Index           =   1
            Left            =   210
            Top             =   465
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   529
            Caption         =   "PNO"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   300
            Index           =   0
            Left            =   645
            Top             =   465
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   529
            Caption         =   "エラー内容"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   375
            Left            =   165
            Top             =   30
            Width           =   14295
            _ExtentX        =   25215
            _ExtentY        =   661
            BackColor       =   0
            ForeColor       =   16711935
            Caption         =   "エラーポケット設定"
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
         Begin vsFlexLib.vsFlexArray vsGrid2 
            Height          =   2070
            Left            =   165
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   420
            Width           =   14295
            _Version        =   196608
            _ExtentX        =   25215
            _ExtentY        =   3651
            _StockProps     =   228
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "PNO|エラー内容        |PNO|エラー内容      |PNO|エラー内容     |PNO|エラー内容     |PNO|エラー内容     |PNO|エラー内容     | "
            Rows            =   6
            Cols            =   13
            FixedCols       =   0
            BackColor       =   16761087
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   0
            ScrollBars      =   0
            RowHeightMin    =   330
            SelectionMode   =   1
            BackColorAlternate=   16761087
         End
      End
   End
   Begin Cis3D_v60.CIS3D PopupSMenu 
      Height          =   1740
      Left            =   8565
      Top             =   5730
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   3069
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
      Begin Cis3D_v60.CIS3D POP_Menu 
         Height          =   285
         Index           =   0
         Left            =   0
         Top             =   15
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   503
         ForeColor       =   0
         Caption         =   "挿入"
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
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D POP_Menu 
         Height          =   285
         Index           =   1
         Left            =   0
         Top             =   300
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   503
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "ｺﾋﾟｰ"
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
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D POP_Menu 
         Height          =   285
         Index           =   2
         Left            =   0
         Top             =   585
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   503
         Caption         =   "貼付け"
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
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D POP_Menu 
         Height          =   285
         Index           =   3
         Left            =   0
         Top             =   870
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   503
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D POP_Menu 
         Height          =   285
         Index           =   4
         Left            =   0
         Top             =   1155
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   503
         Caption         =   "ｸﾘｱ"
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
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D POP_Menu 
         Height          =   285
         Index           =   5
         Left            =   0
         Top             =   1440
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   503
         Caption         =   "戻る"
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
         cAlingnment     =   7
      End
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   345
      ScaleHeight     =   855
      ScaleWidth      =   7590
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1095
      Width           =   7650
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   750
         Left            =   60
         Top             =   60
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1323
         ForeColor       =   192
         Caption         =   " リード グループ"
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
         cPositionY      =   30
      End
      Begin Cis3D_v60.CIS3D OpH1_Group 
         Height          =   750
         Index           =   5
         Left            =   6285
         Top             =   60
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1323
         ForeColor       =   0
         Caption         =   "５"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         cPositionY      =   30
      End
      Begin Cis3D_v60.CIS3D OpH1_Group 
         Height          =   750
         Index           =   1
         Left            =   1125
         Top             =   60
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1323
         ForeColor       =   0
         Caption         =   "１"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         cPositionY      =   30
      End
      Begin Cis3D_v60.CIS3D OpH1_Group 
         Height          =   750
         Index           =   2
         Left            =   2415
         Top             =   60
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1323
         ForeColor       =   0
         Caption         =   "２"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         cPositionY      =   30
      End
      Begin Cis3D_v60.CIS3D OpH1_Group 
         Height          =   750
         Index           =   3
         Left            =   3705
         Top             =   60
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1323
         ForeColor       =   0
         Caption         =   "３"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         cPositionY      =   30
      End
      Begin Cis3D_v60.CIS3D OpH1_Group 
         Height          =   750
         Index           =   4
         Left            =   4995
         Top             =   60
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1323
         ForeColor       =   0
         Caption         =   "４"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         cPositionY      =   30
      End
   End
   Begin Cis3D_v60.CIS3D B1_Area1 
      Height          =   840
      Left            =   315
      Top             =   2235
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   1482
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
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   720
         Left            =   60
         Top             =   60
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1270
         ForeColor       =   16711680
         Caption         =   "振出時間"
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
         Begin Cis3D_v60.CIS3D CIS3D8 
            Height          =   330
            Left            =   525
            Top             =   285
            Width           =   120
            _ExtentX        =   212
            _ExtentY        =   582
            Caption         =   ":"
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
         Begin CisText_V60.CisText iB1_FHh 
            Height          =   360
            Left            =   135
            TabIndex        =   0
            Top             =   285
            Width           =   375
            _ExtentX        =   661
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
         Begin CisText_V60.CisText iB1_FMm 
            Height          =   360
            Left            =   645
            TabIndex        =   1
            Top             =   285
            Width           =   375
            _ExtentX        =   661
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
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   720
         Left            =   1245
         Top             =   60
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   1270
         ForeColor       =   16711680
         Caption         =   "名 称"
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
         Begin CisText_V60.CisText iB1_Name 
            Height          =   360
            Left            =   105
            TabIndex        =   2
            Top             =   285
            Width           =   1890
            _ExtentX        =   3334
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
            Text            =   "XXXXXXXXXXXXXX"
            MaxLength       =   14
            IMEMode         =   4
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
      Left            =   11295
      ScaleHeight     =   660
      ScaleWidth      =   3405
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1080
      Width           =   3465
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   855
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
         PFBackColor     =   255
      End
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   1695
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
         Left            =   2535
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   855
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
      TabIndex        =   9
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
   Begin VB.PictureBox B1_Area3 
      BackColor       =   &H00C0C0C0&
      Height          =   4440
      Left            =   270
      ScaleHeight     =   4380
      ScaleWidth      =   14610
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3735
      Width           =   14670
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   4245
         Left            =   60
         Top             =   60
         Width           =   14505
         _ExtentX        =   25585
         _ExtentY        =   7488
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
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   4185
            Index           =   0
            Left            =   30
            Top             =   30
            Width           =   14445
            _ExtentX        =   25479
            _ExtentY        =   7382
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
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   330
               Index           =   0
               Left            =   120
               Top             =   3735
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   582
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "5"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   330
               Index           =   12
               Left            =   120
               Top             =   3405
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   582
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "4"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   330
               Index           =   13
               Left            =   120
               Top             =   3075
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   582
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "3"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   330
               Index           =   14
               Left            =   120
               Top             =   2745
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   582
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "2"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   330
               Index           =   15
               Left            =   120
               Top             =   2415
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   582
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   315
               Index           =   16
               Left            =   5040
               Top             =   2085
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   556
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "9"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   315
               Index           =   17
               Left            =   2745
               Top             =   2085
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   556
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "8"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   315
               Index           =   18
               Left            =   435
               Top             =   2085
               Width           =   2310
               _ExtentX        =   4075
               _ExtentY        =   556
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "7"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D SaveCol 
               Height          =   315
               Left            =   120
               Top             =   2085
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   556
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "№"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   315
               Index           =   20
               Left            =   7365
               Top             =   2085
               Width           =   2310
               _ExtentX        =   4075
               _ExtentY        =   556
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "10"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   315
               Index           =   21
               Left            =   9675
               Top             =   2085
               Width           =   2310
               _ExtentX        =   4075
               _ExtentY        =   556
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "11"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   315
               Index           =   22
               Left            =   11985
               Top             =   2085
               Width           =   2310
               _ExtentX        =   4075
               _ExtentY        =   556
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "12"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   330
               Index           =   11
               Left            =   120
               Top             =   1755
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   582
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "5"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   330
               Index           =   10
               Left            =   120
               Top             =   1425
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   582
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "4"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   330
               Index           =   9
               Left            =   120
               Top             =   1095
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   582
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "3"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   330
               Index           =   8
               Left            =   120
               Top             =   765
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   582
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "2"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   330
               Index           =   2
               Left            =   120
               Top             =   435
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   582
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   315
               Index           =   4
               Left            =   5040
               Top             =   105
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   556
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "3"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   315
               Index           =   3
               Left            =   2745
               Top             =   105
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   556
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "2"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   315
               Index           =   1
               Left            =   435
               Top             =   105
               Width           =   2310
               _ExtentX        =   4075
               _ExtentY        =   556
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D SaveRow 
               Height          =   315
               Left            =   120
               Top             =   105
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   556
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "№"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D B1_ErrName 
               Height          =   330
               Index           =   5
               Left            =   435
               Top             =   1755
               Visible         =   0   'False
               Width           =   2310
               _ExtentX        =   4075
               _ExtentY        =   582
               BackColor       =   16761087
               ForeColor       =   0
               Caption         =   "あふれ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   1
               cPositionX      =   30
            End
            Begin Cis3D_v60.CIS3D B1_ErrName 
               Height          =   330
               Index           =   4
               Left            =   435
               Top             =   1425
               Visible         =   0   'False
               Width           =   2310
               _ExtentX        =   4075
               _ExtentY        =   582
               BackColor       =   16761087
               ForeColor       =   0
               Caption         =   "抜き取り"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   1
               cPositionX      =   30
            End
            Begin Cis3D_v60.CIS3D B1_ErrName 
               Height          =   330
               Index           =   3
               Left            =   435
               Top             =   1095
               Visible         =   0   'False
               Width           =   2310
               _ExtentX        =   4075
               _ExtentY        =   582
               BackColor       =   16761087
               ForeColor       =   0
               Caption         =   "ﾁｪｯｸｴﾗｰ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   1
               cPositionX      =   30
            End
            Begin Cis3D_v60.CIS3D B1_ErrName 
               Height          =   330
               Index           =   2
               Left            =   435
               Top             =   765
               Visible         =   0   'False
               Width           =   2310
               _ExtentX        =   4075
               _ExtentY        =   582
               BackColor       =   16761087
               ForeColor       =   0
               Caption         =   "重複読取"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   1
               cPositionX      =   30
            End
            Begin Cis3D_v60.CIS3D B1_ErrName 
               Height          =   330
               Index           =   1
               Left            =   435
               Top             =   435
               Visible         =   0   'False
               Width           =   2310
               _ExtentX        =   4075
               _ExtentY        =   582
               BackColor       =   16761087
               ForeColor       =   0
               Caption         =   "読取不能"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   1
               cPositionX      =   30
            End
            Begin Cis3D_v60.CIS3D B1_ErrName 
               Height          =   330
               Index           =   10
               Left            =   2745
               Top             =   1755
               Visible         =   0   'False
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   582
               BackColor       =   16761087
               ForeColor       =   0
               Caption         =   "グループ外"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   1
               cPositionX      =   30
            End
            Begin Cis3D_v60.CIS3D B1_ErrName 
               Height          =   330
               Index           =   9
               Left            =   2745
               Top             =   1425
               Visible         =   0   'False
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   582
               BackColor       =   16761087
               ForeColor       =   0
               Caption         =   "ﾎﾟｹｯﾄｵｰﾊﾞｰ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   1
               cPositionX      =   30
            End
            Begin Cis3D_v60.CIS3D B1_ErrName 
               Height          =   330
               Index           =   8
               Left            =   2745
               Top             =   1095
               Visible         =   0   'False
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   582
               BackColor       =   16761087
               ForeColor       =   0
               Caption         =   "平準化"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   1
               cPositionX      =   30
            End
            Begin Cis3D_v60.CIS3D B1_ErrName 
               Height          =   330
               Index           =   7
               Left            =   2745
               Top             =   765
               Visible         =   0   'False
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   582
               BackColor       =   16761087
               ForeColor       =   0
               Caption         =   "打切(端数)"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   1
               cPositionX      =   30
            End
            Begin Cis3D_v60.CIS3D B1_ErrName 
               Height          =   330
               Index           =   6
               Left            =   2745
               Top             =   435
               Visible         =   0   'False
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   582
               BackColor       =   16761087
               ForeColor       =   0
               Caption         =   "打切(完了)"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   1
               cPositionX      =   30
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   315
               Index           =   5
               Left            =   7365
               Top             =   105
               Width           =   2310
               _ExtentX        =   4075
               _ExtentY        =   556
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "4"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   315
               Index           =   6
               Left            =   9675
               Top             =   105
               Width           =   2310
               _ExtentX        =   4075
               _ExtentY        =   556
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "5"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   315
               Index           =   7
               Left            =   11985
               Top             =   105
               Width           =   2310
               _ExtentX        =   4075
               _ExtentY        =   556
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "6"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin vsFlexLib.vsFlexArray vsGrid 
               Height          =   4035
               Left            =   75
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   75
               Width           =   14265
               _Version        =   196608
               _ExtentX        =   25162
               _ExtentY        =   7117
               _StockProps     =   228
               Appearance      =   1
               ConvInfo        =   1418783674
               FormatString    =   "NO|XXXXXXX|XX|XXXXXXXX|XXXXXXX|XX|XXXXXXXX|XXXXXXX|XX|XXXXXXXX|XXXXXXX|XX|XXXXXXXX|XXXXXXX|XX|XXXXXXXX|XXXXXXX|XX|XXXXXXXX|"
               Rows            =   12
               Cols            =   20
               BackColor       =   12648384
               BackColorBkg    =   12632256
               FocusRect       =   0
               HighLight       =   0
               RowHeightMin    =   330
               SelectionMode   =   1
               Begin CisText_V60.CisText iB1_Torcd 
                  Height          =   375
                  Index           =   1
                  Left            =   330
                  TabIndex        =   14
                  Top             =   315
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   661
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
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
               Begin CisText_V60.CisText iB1_Ukeir 
                  Height          =   375
                  Index           =   1
                  Left            =   1245
                  TabIndex        =   15
                  Top             =   315
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   661
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
               Begin CisText_V60.CisText iB1_Torcd 
                  Height          =   375
                  Index           =   2
                  Left            =   2625
                  TabIndex        =   16
                  Top             =   315
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   661
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
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
               Begin CisText_V60.CisText iB1_Ukeir 
                  Height          =   375
                  Index           =   2
                  Left            =   3540
                  TabIndex        =   17
                  Top             =   315
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   661
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
               Begin CisText_V60.CisText iB1_Torcd 
                  Height          =   375
                  Index           =   3
                  Left            =   4950
                  TabIndex        =   18
                  Top             =   300
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   661
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
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
               Begin CisText_V60.CisText iB1_Ukeir 
                  Height          =   375
                  Index           =   3
                  Left            =   5865
                  TabIndex        =   19
                  Top             =   300
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   661
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
               Begin CisText_V60.CisText iB1_Torcd 
                  Height          =   375
                  Index           =   4
                  Left            =   7245
                  TabIndex        =   20
                  Top             =   300
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   661
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
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
               Begin CisText_V60.CisText iB1_Ukeir 
                  Height          =   375
                  Index           =   4
                  Left            =   8160
                  TabIndex        =   21
                  Top             =   300
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   661
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
               Begin CisText_V60.CisText iB1_Torcd 
                  Height          =   375
                  Index           =   5
                  Left            =   9570
                  TabIndex        =   22
                  Top             =   315
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   661
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
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
               Begin CisText_V60.CisText iB1_Ukeir 
                  Height          =   375
                  Index           =   5
                  Left            =   10485
                  TabIndex        =   23
                  Top             =   315
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   661
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
               Begin CisText_V60.CisText iB1_Torcd 
                  Height          =   375
                  Index           =   6
                  Left            =   11865
                  TabIndex        =   24
                  Top             =   315
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   661
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
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
               Begin CisText_V60.CisText iB1_Ukeir 
                  Height          =   375
                  Index           =   6
                  Left            =   12780
                  TabIndex        =   25
                  Top             =   315
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   661
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
            End
            Begin CisText_V60.CisText iB1_DummyED 
               Height          =   375
               Left            =   12345
               TabIndex        =   26
               Top             =   4170
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   661
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
            Begin CisText_V60.CisText iB1_DummyST 
               Height          =   375
               Left            =   11715
               TabIndex        =   11
               Top             =   4215
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   661
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
         End
      End
   End
   Begin Cis3D_v60.CIS3D B1_Area2 
      Height          =   855
      Left            =   3765
      Top             =   2220
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   1508
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
         Height          =   720
         Left            =   660
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1270
         ForeColor       =   16711680
         Caption         =   "1段"
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
         Begin CisText_V60.CisText iB1_Maisu 
            Height          =   360
            Index           =   1
            Left            =   150
            TabIndex        =   3
            Top             =   285
            Width           =   510
            _ExtentX        =   900
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
      Begin Cis3D_v60.CIS3D CIS3D10 
         Height          =   720
         Left            =   1485
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1270
         ForeColor       =   16711680
         Caption         =   "2段"
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
         Begin CisText_V60.CisText iB1_Maisu 
            Height          =   360
            Index           =   2
            Left            =   150
            TabIndex        =   4
            Top             =   285
            Width           =   510
            _ExtentX        =   900
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
      Begin Cis3D_v60.CIS3D CIS3D12 
         Height          =   720
         Left            =   2310
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1270
         ForeColor       =   16711680
         Caption         =   "3段"
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
         Begin CisText_V60.CisText iB1_Maisu 
            Height          =   360
            Index           =   3
            Left            =   150
            TabIndex        =   5
            Top             =   285
            Width           =   510
            _ExtentX        =   900
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
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   720
         Left            =   3135
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1270
         ForeColor       =   16711680
         Caption         =   "4段"
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
         Begin CisText_V60.CisText iB1_Maisu 
            Height          =   360
            Index           =   4
            Left            =   150
            TabIndex        =   6
            Top             =   285
            Width           =   510
            _ExtentX        =   900
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
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   720
         Left            =   3960
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1270
         ForeColor       =   16711680
         Caption         =   "5段"
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
         Begin CisText_V60.CisText iB1_Maisu 
            Height          =   360
            Index           =   5
            Left            =   150
            TabIndex        =   7
            Top             =   285
            Width           =   510
            _ExtentX        =   900
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
      Begin Cis3D_v60.CIS3D CIS3D15 
         Height          =   720
         Left            =   60
         Top             =   75
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1270
         ForeColor       =   0
         Caption         =   "設定枚数"
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
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 かんばんポケット指定  】"
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
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14370
         Top             =   75
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
         Left            =   11850
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
   End
End
Attribute VB_Name = "CKN0050"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   かんばんポケット指定
'**       フォームID    :   CKN0050
'**       処理概要      :
'**
'**       作  成  日    :   2003/11/19  By CIS SIMPR-A
'**       変  更  日    :   2006/01/10  エラーポケットの可変対応
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    Dim CisVsGrid2      As New CisVsGrid3
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim ActObj          As Object
    Dim ActRow          As Integer
    
    Dim StartFlg        As Boolean
    Dim EndFlg          As Boolean
    Dim mCisMenu        As CIS3D
    Dim CopyTorcd       As String       ' 仕入先コピー
    Dim CopyUkeir       As String       ' 受入コピー
    Dim CopyTornm       As String       ' 仕入先名コピー
    Dim CopyRow       As Integer       ' 仕入先コピー
    Dim CopyCol       As Integer       ' 受入コピー
    Dim mSaveRow       As Integer       ' Row退避
    Dim mSaveCol       As Integer       ' Col退避

    Dim mMColor         As Long
    Dim mPocketMax      As Long
    Dim mEColor         As Long
    Dim mEPocket(60)    As Byte
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
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止
        
    ' 名称マスタ確認
    Call MeisyoInit
    '+-----------------------+
    '+  グリッド初期設定
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
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid2
        Set .GridObj = vsGrid2
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
'        .SelectCol = 19
        .InitGet = False
        .Init
    End With
    Call GridInit2
    mEColor = vsGrid2.BackColor
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "B1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = ""
    Call SyoriName(SyoriKB)
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
'初期表示
    Call GridInitDisp
    Call OpH1_Group_Click(1)
'
    FormAct = Not FormAct
'
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
           Case vbKeyF1:       Call OpH1_Group_Click(1)  '【GROUP1】
           Case vbKeyF2:       Call OpH1_Group_Click(2)  '【GROUP2】
           Case vbKeyF3:       Call OpH1_Group_Click(3)  '【GROUP3】
           Case vbKeyF4:       Call OpH1_Group_Click(4)  '【GROUP4】
           Case vbKeyF5:       Call OpH1_Group_Click(5)  '【GROUP5】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyF9:       If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【削除取消】
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
'+    削除取消キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    SyoriKB = ""
    Call SyoriName(SyoriKB)
End Sub
'+----------------------------+
'+    GROUP削除キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
        SyoriKB = "DEL"
    Else
        SyoriKB = ""
    End If
    Call SyoriName(SyoriKB)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iB1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iB1_Ukeir" Then Call Look_Cycle: Exit Sub
End Sub
'+----------------------------+
'+       取引先検索            +
'+----------------------------+
Private Sub Look_Torcd()

    RV_Left = 0: RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 8
'
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
       iB1_Torcd(CisFun.Val2(iB1_Torcd(1).Tag)) = RTrim(RV_TorcdK)
       Call iB1_Torcd_LostFocus(CisFun.Val2(iB1_Torcd(1).Tag))
    End If

End Sub
'+----------------------------+
'+       受入検索            +
'+----------------------------+
Private Sub Look_Ukeir()

    RV_Left = 0: RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 8
    RV_Torcd = iB1_Torcd(CisFun.Val2(SaveCol.Tag))
    RV_Kojyo = ""
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
       iB1_Torcd(CisFun.Val2(SaveCol.Tag)) = RTrim(RV_TorcdK)
       Call iB1_Torcd_LostFocus(CisFun.Val2(SaveCol.Tag))
       iB1_Ukeir(CisFun.Val2(SaveCol.Tag)) = RTrim(RV_Ukeir)
    End If

End Sub
'+----------------------------+
'+       発注サイクル検索      +
'+----------------------------+
Private Sub Look_Cycle()

    RV_Left = 0:    RV_Top = 0
    RV_Call = "C"
    RV_Torcd = "": RV_Ukeir = "": RV_Kojyo = ""
    RV_TorcdK = iB1_Torcd(CisFun.Val2(SaveCol.Tag))
    
    CKNK0100.Show vbModal
    Unload CKNK0100
    Set CKNK0100 = Nothing

    If RV_Rtn Then
       iB1_Torcd(CisFun.Val2(SaveCol.Tag)) = RTrim(RV_Torcd)
       Call iB1_Torcd_LostFocus(CisFun.Val2(SaveCol.Tag))
       iB1_Ukeir(CisFun.Val2(SaveCol.Tag)) = RTrim(RV_Ukeir)
    End If

End Sub
'*============================================================'
'*      グループ選択
'*============================================================'
Private Sub OpH1_Group_Click(Index As Integer)
    If Index = CisFun.Val2(OpH1_Group(1).Tag) Then Exit Sub
    Call BodyDisp(Index)
    For gInt = 1 To 5
        OpH1_Group(gInt).BackColor = gPLostSel
    Next
    OpH1_Group(Index).BackColor = gPGotSel
    OpH1_Group(1).Tag = Format(Index, "0")
    'Grid １行目表示
    Call VsGridToText(1)

End Sub
'+-----------------------------+
'+  iB1_FHh
'+-----------------------------+
Private Sub iB1_FHh_GotFocus()
    EndFlg = False
    Set ActObj = iB1_FHh
End Sub
Private Sub iB1_FHh_LostFocus()
    iB1_FHh = Format(CisFun.Val2(iB1_FHh), "00")
End Sub
Private Sub iB1_FMm_LostFocus()
    iB1_FMm = Format(CisFun.Val2(iB1_FMm), "00")
End Sub
'+-----------------------------+
'+  iB1_Maisu
'+-----------------------------+
Private Sub iB1_Maisu_GotFocus(Index As Integer)
    If Index = 5 Then StartFlg = False
End Sub
'+-----------------------------+
'+  iB1_Torcd
'+-----------------------------+
Private Sub iB1_Torcd_GotFocus(Index As Integer)
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    iB1_Torcd(1).Tag = Format(Index, "00")
    Set ActObj = iB1_Torcd(Index)
    SaveCol.Tag = Format(Index, "0")
    If Index = 1 Or Index = 3 Then StartFlg = True
End Sub
Private Sub iB1_Torcd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Torcd_LostFocus(Index As Integer)
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    vsGrid.TextMatrix(CisFun.Val2(SaveRow.Tag), (Index - 1) * 3 + 2) = ""
    vsGrid.TextMatrix(CisFun.Val2(SaveRow.Tag), (Index - 1) * 3 + 3) = ""
    If RTrim(iB1_Torcd(Index)) <> "" Then
        If TorNmGet(iB1_Torcd(Index), 8) Then
            vsGrid.TextMatrix(CisFun.Val2(SaveRow.Tag), (Index - 1) * 3 + 3) = CisFun.Left2(TRM.略称, 10)
        End If
    Else
        iB1_Ukeir(Index) = ""
    End If
    With vsGrid
        ActRow = CisFun.Val2(SaveRow)
        .TextMatrix(ActRow, (Index - 1) * 3 + 1) = RTrim(iB1_Torcd(Index))
        .TextMatrix(ActRow, (Index - 1) * 3 + 2) = RTrim(iB1_Ukeir(Index))
    End With
End Sub
'+-----------------------------+
'+  iB1_Ukeir
'+-----------------------------+
Private Sub iB1_Ukeir_GotFocus(Index As Integer)
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Set ActObj = iB1_Ukeir(Index)
    SaveCol.Tag = Format(Index, "0")
    If Index = 6 Then EndFlg = True
End Sub
Private Sub iB1_Ukeir_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Ukeir_LostFocus(Index As Integer)
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    iB1_Ukeir(1).Tag = Format(Index, "00")
    With vsGrid
        ActRow = CisFun.Val2(SaveRow)
        .TextMatrix(ActRow, (Index - 1) * 3 + 2) = RTrim(iB1_Ukeir(Index))
    End With
End Sub
Private Sub iB1_DummyED_GotFocus()
    If CisFun.Val2(SaveRow.Tag) < 5 Then
'        iB1_Torcd(3).SetFocus
        iB1_Torcd(1).SetFocus
    Else
        If CisFun.Val2(SaveRow.Tag) = 11 Then
            SendKeys "{tab}"
            Exit Sub
        End If
        iB1_Torcd(1).SetFocus
    End If
    If Not EndFlg Then Exit Sub
    Call Grid_RowChange(vbKeyDown)
End Sub
Private Sub iB1_DummyST_GotFocus()
    If CisFun.Val2(SaveRow.Tag) = 1 Then
        SendKeys "{tab}"
        Exit Sub
    End If
'''    If CisFun.Val2(SaveRow.Tag) > 6 Then
        iB1_Torcd(1).SetFocus
'''    Else
'''        iB1_Torcd(3).SetFocus
'''    End If
    If Not StartFlg Then Exit Sub
    Call Grid_RowChange(vbKeyUp)
End Sub
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   Dim W_Color     As Long
   W_Color = &H80000008 '初期色ｾｯﾄ(黒)
   Select Case SyoriKB
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
      PB_CAN.Visible = False
      Exit Sub
    Else
      PB_CAN.Visible = True
   End If
   K_Sykbnm.Caption = SyoriNM
   K_Sykbnm.Visible = True
   Timer2.Enabled = True
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
    H1Mode = False
    H1Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
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
    End Select
    H1_Area1.Enabled = B1Mode
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B1_Area2.Enabled = B1Mode
    B1_Area2.BackColor = B1Color
    B1_Area3.Enabled = B1Mode
    B1_Area3.BackColor = B1Color
    PB_Look.Visible = H1Mode
    
    Call GridTextEnabled(B1Mode)
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
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If ClrType = "ALL" Or ClrType = "VS" Then
        For gCnt1 = 1 To 11
            For gCnt2 = 1 To 18
                vsGrid.TextMatrix(gCnt1, gCnt2) = ""
            Next
        Next
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
' << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call BodyDisp(CisFun.Val2(OpH1_Group(1).Tag))
            Call GridTextEnabled(True)
            GoTo ReturnPress_Ed
        End If
        If Not DBPut Then GoTo ReturnPress_Ed
        vsGrid.Row = 1
        Call VsGridToText(vsGrid.Row)
        ProcHB = "B1"
        Call DispChange(ProcHB)
'        Call SyoriName(SyoriKB)
        GoTo ReturnPress_Ed
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    
    Body1Chk = False
    Call TextToVsGrid(Val(SaveRow.Tag))
    vsGrid.SetFocus
'
    For gInt = 1 To 60
        'Grid Row,Col
        Call RowCol_Set(gInt)
    
        If RTrim(vsGrid.TextMatrix(mSaveRow, (mSaveCol - 1) * 3 + 1)) <> "" Then
            'エラーポケットへの入力、最大数以上への入力チェック     2006.01.10追加
            If mEPocket(gInt) <> 0 Then
                With CisFun
                    .MB_Lines = 4
                    If mEPocket(gInt) = 1 Then
                        .MB_MSG(1) = "    【　エラーポケットへの入力不可！　】            "
                    Else
                        .MB_MSG(1) = "    【　ポケット最大数以降への入力不可！　】            "
                    End If
                    .MB_MSG(3) = "    仕入先・受入をクリアして下さい            "
                            
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                iB1_Torcd(mSaveCol).SetFocus
                GoTo Body1Chk_Err
            End If
            gSL_Select = "SELECT 取引先 FROM 取引先マスタ"
            gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
            gSL_Select = gSL_Select & "   AND 取引先CD = '" & RTrim(vsGrid.TextMatrix(mSaveRow, (mSaveCol - 1) * 3 + 1)) & "'"
            If Not TRMRead(gSL_Select, 1) Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(1) = "    【　仕入先　未登録！　】            "
                    .MB_MSG(3) = "    仕入先を再入力して下さい            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                iB1_Torcd(mSaveCol).SetFocus
                GoTo Body1Chk_Err
            End If
            gSL_Select = "SELECT COUNT(仕入先) 件数 FROM 発注サイクルマスタ"
            gSL_Select = gSL_Select & " WHERE 仕入先 = '" & RTrim(vsGrid.TextMatrix(mSaveRow, (mSaveCol - 1) * 3 + 1)) & "'"
            If HCMRead(gSL_Select, 1) Then
                If HCM.件数 = 0 Then
                    With CisFun
                        .MB_Lines = 4
                        .MB_MSG(1) = "    【　発注サイクル　未登録！　】            "
                        .MB_MSG(3) = "    仕入先を再入力して下さい            "
                        .MB_Title = ""
                        .MB_Button = Error
                        .MBOX
                    End With
                    iB1_Torcd(mSaveCol).SetFocus
                    GoTo Body1Chk_Err
                End If
            End If
            If RTrim(vsGrid.TextMatrix(mSaveRow, (mSaveCol - 1) * 3 + 2)) <> "" Then
                gSL_Select = "SELECT 受入 FROM 取引先受入マスタ"
                gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
                gSL_Select = gSL_Select & " AND 取引先CD = '" & RTrim(vsGrid.TextMatrix(mSaveRow, (mSaveCol - 1) * 3 + 1)) & "'"
                gSL_Select = gSL_Select & " AND 受入 = '" & RTrim(vsGrid.TextMatrix(mSaveRow, (mSaveCol - 1) * 3 + 2)) & "'"
                If Not TUMRead(gSL_Select, 1) Then
                    With CisFun
                        .MB_Lines = 4
                        .MB_MSG(1) = "    【　受入　未登録！　】            "
                        .MB_MSG(3) = "    受入を再入力して下さい            "
                        .MB_Title = ""
                        .MB_Button = Error
                        .MBOX
                    End With
                    iB1_Ukeir(mSaveCol).SetFocus
                    GoTo Body1Chk_Err
                End If
            End If
        End If
    Next
    
    Call GridTextEnabled(False)
    Body1Chk = True
    Exit Function
Body1Chk_Err:
    vsGrid.Row = mSaveRow
    Call VsGridToText(vsGrid.Row)
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
'**************************************
'*          BODY 内容ｾｯﾄ処理           *
'**************************************
Private Sub BodyDisp(SetGrp As Integer)
    Static wStr()  As String
    Call HeadBodyClear("B")
'かんばん読取エラーポケット表示
    Erase wStr
    gSL_Select = " SELECT * FROM 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = 'かんばんポケット内容'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   And 値 = " & SetGrp
    If SYMRead(gSL_Select, 0) Then
        iB1_FHh = RTrim(Left(SYM.英数字1, 2))
        iB1_FMm = RTrim(Mid(SYM.英数字1, 3, 2))
        iB1_Name = RTrim(SYM.英数字2)
        wStr = Split(SYM.英数字3, ",")
        For gInt = 0 To UBound(wStr)
            iB1_Maisu(gInt + 1) = wStr(gInt)
        Next
    End If

'取引先・受入表示
    Call HeadBodyClear("VS")
    
    gSL_Select = " SELECT KP.取引先,KP.受入,ポケットNO,ISNULL(TR.略称,'')取引先名"
    gSL_Select = gSL_Select & " FROM かんばんポケットテーブル KP"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先ＣＤ = KP.取引先"
    gSL_Select = gSL_Select & " WHERE グループ = " & SetGrp
    gSL_Select = gSL_Select & " ORDER BY ポケットNO"
    If KPTRead(gSL_Select, 0) Then
        Call GridSet
    Else
        vsGrid.Row = 1: vsGrid.Col = 7
    End If
    Call KPTClose

    
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
        If Not .MBOX Then
            Call VsGridToText(CisFun.Val2(SaveRow.Tag))
            Call GridTextEnabled(True)
            Exit Function
        End If
    End With
    
    With CisDB
        On Error GoTo DBDelete_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 削除処理 )
        ' 取引先マスタ
        .SQL = "DELETE FROM かんばんポケットテーブル "
        .SQL = .SQL & " WHERE グループ = " & CisFun.Val2(OpH1_Group(1).Tag) & ""
        .DBExec
        ' 取引先受入マスタ
        .SQL = "UPDATE 名称マスタ SET"
        .SQL = .SQL & " 英数字1 = '',英数字2 = '',英数字3 = ''"
        .SQL = .SQL & " WHERE 区分名称 = 'かんばんポケット内容'"
        .SQL = .SQL & "   AND 区分タイプ = ''"
        .SQL = .SQL & "   AND 値 = " & RTrim(OpH1_Group(1).Tag) & ""
        On Error GoTo DBDelete_Err
        .DBExec
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

'        iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")

        On Error GoTo 0
    End With
    DBDelete = True
    Exit Function
DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False
    Dim AddFlg  As Byte

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     更新処理を行います      "
        .MB_Title = "更新処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Call VsGridToText(CisFun.Val2(SaveRow.Tag))
            Call GridTextEnabled(True)
            ActObj.SetFocus
            Exit Function
        End If
    End With
    On Error GoTo DBPut_Err

    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
'かんばんポケットテーブル更新
    For gInt = 1 To 60
        'Grid Row,Col
        Call RowCol_Set(gInt)
        
        gSL_Select = "SELECT * From かんばんポケットテーブル"
        gSL_Select = gSL_Select & " WHERE グループ = " & CisFun.Val2(OpH1_Group(1).Tag) & ""
        gSL_Select = gSL_Select & "   And ポケットNO = " & gInt & ""
        If KPTRead(gSL_Select, 1) Then
            AddFlg = 1
            If RTrim(KPT.取引先) = RTrim(vsGrid.TextMatrix(mSaveRow, (mSaveCol - 1) * 3 + 1)) And _
                RTrim(KPT.受入) = RTrim(vsGrid.TextMatrix(mSaveRow, (mSaveCol - 1) * 3 + 2)) Then
                AddFlg = 2
            End If
            If RTrim(vsGrid.TextMatrix(mSaveRow, (mSaveCol - 1) * 3 + 1)) = "" And _
                RTrim(vsGrid.TextMatrix(mSaveRow, (mSaveCol - 1) * 3 + 2)) = "" Then
                AddFlg = 3
            End If
        Else
            If RTrim(vsGrid.TextMatrix(mSaveRow, (mSaveCol - 1) * 3 + 1)) <> "" Then
                Call ItemsClearKPT
                AddFlg = 0
            Else
                AddFlg = 2
            End If
        End If
        KPT.グループ = CisFun.Val2(OpH1_Group(1).Tag)
        KPT.ポケットNO = gInt
        KPT.取引先 = RTrim(vsGrid.TextMatrix(mSaveRow, (mSaveCol - 1) * 3 + 1))
        KPT.受入 = RTrim(vsGrid.TextMatrix(mSaveRow, (mSaveCol - 1) * 3 + 2))
'
    ' ( 追加処理 )
        If AddFlg = 0 Then
            KPT.作成者 = gTanto
            Call KPTInsert
        Else
    ' ( 更新処理 )
            If AddFlg = 1 Then
                KPT.更新者 = gTanto
                Call KPTUpdate
            End If
    ' ( 削除処理 )
            If AddFlg = 3 Then
                With CisDB
                    .SQL = "DELETE FROM かんばんポケットテーブル "
                    .SQL = .SQL & " WHERE グループ = " & CisFun.Val2(OpH1_Group(1).Tag) & ""
                    .SQL = .SQL & "   AND ポケットNO = " & KPT.ポケットNO & ""
                    .DBExec
                End With
            End If
        End If
'
    Next gInt
'名称マスタ更新（かんばんポケット内容）
    gSL_Select = "SELECT * From 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = 'かんばんポケット内容'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND 値 = " & CisFun.Val2(OpH1_Group(1).Tag) & ""
    If SYMRead(gSL_Select, 1) Then
        AddFlg = 1
    Else
        AddFlg = 0
    End If
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "かんばんポケット内容"
        .区分桁数 = 1
        .値 = RTrim(OpH1_Group(1).Tag)
        .値名称 = "グループ" & RTrim(OpH1_Group(1).Tag)
        .英数字1 = iB1_FHh & iB1_FMm
        .英数字2 = RTrim(iB1_Name)
        .英数字3 = CisFun.Val2(iB1_Maisu(1)) & "," & CisFun.Val2(iB1_Maisu(2)) & ","
        .英数字3 = .英数字3 & CisFun.Val2(iB1_Maisu(3)) & "," & CisFun.Val2(iB1_Maisu(4)) & ","
        .英数字3 = .英数字3 & CisFun.Val2(iB1_Maisu(5))
        .作成者 = gTanto
        .更新者 = gTanto
        .制御区分 = "X"
    End With
    If AddFlg = 0 Then
        Call SYMInitCreate
    Else
        With CisDB
            .SQL = "UPDATE 名称マスタ SET "
            .SQL = .SQL & " 英数字1 = '" & SYMCreate.英数字1 & "',"
            .SQL = .SQL & " 英数字2 = '" & SYMCreate.英数字2 & "',"
            .SQL = .SQL & " 英数字3 = '" & SYMCreate.英数字3 & "',"
            .SQL = .SQL & " 更新者 = '" & SYMCreate.更新者 & "',"
            .SQL = .SQL & " 更新日 = GETDATE(),"
            .SQL = .SQL & " 更新端末 = HOST_NAME()"
            .SQL = .SQL & " WHERE 区分名称 = 'かんばんポケット内容'"
            .SQL = .SQL & "   AND 区分タイプ = ''"
            .SQL = .SQL & "   AND 値 = '" & SYMCreate.値 & "'"
            .DBExec
    
            On Error GoTo 0
        End With
    
    End If
    
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+---------------------------------------+
'+      名称マスタ初期設定
'+---------------------------------------+
Private Sub MeisyoInit()

' かんばん読取エラーポケット
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "かんばんエラーポケット"
        .区分桁数 = 2
        .値 = "01"
        .値名称 = "読取不能"
        .作成者 = gTanto
        .制御区分 = "X"
    End With
    If Not SYMInitCreate(True) Then
        For gInt = 2 To 10
            With SYMCreate
                .値 = Format(gInt, "00")
                Select Case gInt
                     Case 2: .値名称 = "重複読取"
                     Case 3: .値名称 = "ﾁｪｯｸｴﾗｰ"
                     Case 4: .値名称 = "抜き取り"
                     Case 5: .値名称 = "あふれ"
                     Case 6: .値名称 = "打切(完了)"
                     Case 7: .値名称 = "打切(端数)"
                     Case 8: .値名称 = "平準化"
                     Case 9: .値名称 = "ﾎﾟｹｯﾄｵｰﾊﾞｰ"
                     Case 10: .値名称 = "グループ外"
                              .数字1 = 1
                End Select
            End With
            Call SYMInitCreate
        Next gInt
    End If
' かんばんポケット内容
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "かんばんポケット内容"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "グループ１"
        .英数字3 = "70,70,70,120,250"
        .作成者 = gTanto
        .制御区分 = "X"
    End With
    If Not SYMInitCreate(True) Then
        For gInt = 2 To 5
            With SYMCreate
                .値 = Format(gInt, "#0")
                Select Case gInt
                     Case 2: .値名称 = "グループ２"
                     Case 3: .値名称 = "グループ３"
                     Case 4: .値名称 = "グループ４"
                     Case 5: .値名称 = "グループ５"
                End Select
            End With
            Call SYMInitCreate
        Next gInt
    End If
End Sub
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 2, 0, "NO")
        Call .FixedSet(中中, 左中, あり, 7, 1, "取引先1")
        Call .FixedSet(中中, 中中, あり, 2, 2, "受入1")
        Call .FixedSet(中中, 左中, あり, 8, 3, "取引先名1")
        Call .FixedSet(中中, 左中, あり, 7, 4, "取引先2")
        Call .FixedSet(中中, 中中, あり, 2, 5, "受入2")
        Call .FixedSet(中中, 左中, あり, 8, 6, "取引先名2")
        Call .FixedSet(中中, 左中, あり, 7, 7, "取引先3")
        Call .FixedSet(中中, 中中, あり, 2, 8, "受入3")
        Call .FixedSet(中中, 左中, あり, 8, 9, "取引先名3")
        Call .FixedSet(中中, 左中, あり, 7, 10, "取引先4")
        Call .FixedSet(中中, 中中, あり, 2, 11, "受入4")
        Call .FixedSet(中中, 左中, あり, 8, 12, "取引先名4")
        Call .FixedSet(中中, 左中, あり, 7, 13, "取引先5")
        Call .FixedSet(中中, 中中, あり, 2, 14, "受入5")
        Call .FixedSet(中中, 左中, あり, 8, 15, "取引先名5")
        Call .FixedSet(中中, 左中, あり, 7, 16, "取引先6")
        Call .FixedSet(中中, 中中, あり, 2, 17, "受入6")
        Call .FixedSet(中中, 左中, あり, 8, 18, "取引先名6")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 19, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18")
        .InitDisp
    End With
    
    For gCnt1 = 1 To 11
        With vsGrid
            For gCnt2 = 1 To 6
                .Row = gCnt1: .Col = gCnt2 * 3
                .CellBackColor = &HFFFFC0
                .CellFontSize = 9
            Next gCnt2
        End With
    Next gCnt1
    
End Sub
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub

'テキストからグリッドに退避
    Call TextToVsGrid(Val(SaveRow.Tag))
    gInt = vsGrid.Row
    '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        If gInt - 1 = 6 Then gInt = gInt - 1
        gInt = gInt - 1
    End If
    '↓
    If KeyCode = vbKeyDown Then
        If gInt + 1 > vsGrid.Rows - 1 Then
            GoTo Grid_RowChange_End
        End If
        If gInt + 1 = 6 Then gInt = gInt + 1
        gInt = gInt + 1
    End If
    'PageUp
'    If KeyCode = vbKeyPageUp Then
'        If gInt - CisVsGrid.DispMax < 1 Then
'            gInt = 1
'        Else
'            gInt = gInt - CisVsGrid.DispMax
'        End If
'    End If
    'PageDown
'    If KeyCode = vbKeyPageDown Then
'        If gInt + CisVsGrid.DispMax > Ttl_Kensu - 1 Then
'            gInt = Ttl_Kensu - 1
'        Else
'            gInt = gInt + CisVsGrid.DispMax
'        End If
'    End If
    vsGrid.Row = gInt
'グリッドからテキストに表示
    Call VsGridToText(vsGrid.Row)
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキスト有効・無効
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
    For gInt = 1 To 6
        iB1_Torcd(gInt).Visible = EnabledFlg
        iB1_Ukeir(gInt).Visible = EnabledFlg
    Next
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    With vsGrid
        For gCnt1 = 1 To 6
            .TextMatrix(tRow, (gCnt1 - 1) * 3 + 1) = iB1_Torcd(gCnt1)
            If RTrim(iB1_Torcd(gCnt1)) <> "" Then
                .TextMatrix(tRow, (gCnt1 - 1) * 3 + 2) = iB1_Ukeir(gCnt1)
            Else
                .TextMatrix(tRow, (gCnt1 - 1) * 3 + 2) = ""
            End If
            If TorNmGet(iB1_Torcd(gCnt1), 8) Then
                .TextMatrix(tRow, (gCnt1 - 1) * 3 + 3) = CisFun.Left2(TRM.略称, 10)
            Else
                .TextMatrix(tRow, (gCnt1 - 1) * 3 + 3) = ""
            End If
        Next gCnt1
    End With
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With vsGrid
        SaveRow.Tag = tRow
        For gInt = 1 To 6
            '取引先
            iB1_Torcd(gInt) = RTrim(.TextMatrix(tRow, (gInt - 1) * 3 + 1))
            iB1_Torcd(gInt).Move iB1_Torcd(gInt).Left, vsGrid.CellTop, iB1_Torcd(gInt).Width, vsGrid.CellHeight
            '受入
            iB1_Ukeir(gInt) = RTrim(.TextMatrix(tRow, (gInt - 1) * 3 + 2))
            iB1_Ukeir(gInt).Move iB1_Ukeir(gInt).Left, vsGrid.CellTop, iB1_Ukeir(gInt).Width, vsGrid.CellHeight
        Next
    End With
''''タブ制御    2006.01.10取消
'''    If tRow > 0 And tRow < 6 Then
'''        iB1_Torcd(1).TabStop = False
'''        iB1_Torcd(2).TabStop = False
'''        iB1_Ukeir(1).TabStop = False
'''        iB1_Ukeir(2).TabStop = False
'''        If CisFun.Val2(SaveCol.Tag) <= 2 Then
'''            Set ActObj = iB1_Torcd(3)
'''            ActObj.SetFocus
'''        End If
'''    Else
'''        iB1_Torcd(1).TabStop = True
'''        iB1_Torcd(2).TabStop = True
'''        iB1_Ukeir(1).TabStop = True
'''        iB1_Ukeir(2).TabStop = True
'''    End If
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(CisFun.Val2(ID))
        Next ID
    End With
    ID = 0
    Do Until Not KPT_RDSTS
        'Grid Row,Col
        Call RowCol_Set(CInt(KPT.ポケットNO))
        
        vsGrid.TextMatrix(mSaveRow, (mSaveCol - 1) * 3 + 1) = RTrim(KPT.取引先)
        vsGrid.TextMatrix(mSaveRow, (mSaveCol - 1) * 3 + 2) = RTrim(KPT.受入)
        vsGrid.TextMatrix(mSaveRow, (mSaveCol - 1) * 3 + 3) = RTrim(CisFun.Left2(KPT.取引先名, 10))
        Call KPTReadNext
    Loop

    With vsGrid
         .Row = 1: .Col = 7: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++     ﾃｷｽﾄ入力時      +++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub VSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown Then Exit Sub
    Call TextToVsGrid(Val(SaveRow.Tag))
    Call VsGridToText(vsGrid.Row)
End Sub
Private Sub VSGrid_Click()
    Call TextToVsGrid(Val(SaveRow.Tag))
    Call VsGridToText(vsGrid.Row)
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub
Private Sub vsGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsGrid.MouseRow = 0 Then Exit Sub
    If Button = 2 Then
'        '入力項目非表示
        vsGrid.SetFocus
        Call GridTextEnabled(False)
        CopyRow = vsGrid.MouseRow: CopyCol = vsGrid.MouseCol
        If CopyCol Mod 3 = 0 Then
            CopyCol = CisFun.HasuProc(CopyCol / 3, 切捨て)
        Else
            CopyCol = CisFun.HasuProc(CopyCol / 3, 切捨て) + 1
        End If
        
'        Pop_Row.Tag = .Row
        'ﾎﾟｯﾌﾟｱｯﾌﾟﾒﾆｭｰ表示
        Call PopupMenu_Visible(vsGrid.MouseCol, X, Y)
    End If
End Sub
'+----------------------------------------------------------+
'+      ポップアップメニューの表示
'+----------------------------------------------------------+
'《 MouseMove イベント》
Private Sub PopupMenu_Visible(pCol As Integer, XLeft As Single, YTop As Single)
    
    '選択項目クリア
    For Each mCisMenu In POP_Menu
        mCisMenu.cBoderWidth = 1
        mCisMenu.BackColor = &HC0C0C0
    Next mCisMenu
'    '処理の選択可否
    '貼り付け可否
    If RTrim(CopyTorcd) = "" And RTrim(CopyUkeir) = "" Then
        POP_Menu(2).Enabled = False
        POP_Menu(2).ForeColor = &HE0E0E0
    Else
        POP_Menu(2).Enabled = True
        POP_Menu(2).ForeColor = &H80000008
    End If
    '表示位置
    PopupSMenu.Left = XLeft + 270 + 165 + 60
    PopupSMenu.Top = YTop + 3735 + 165 + 60
    '表示
    PopupSMenu.Visible = True
End Sub
'+----------------------------------------------------------+
'+      cmd_Exec(実行ﾓｼﾞｭｰﾙボタン)
'+----------------------------------------------------------+
'《 MouseMove イベント》
Private Sub Pop_Menu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Mouse_Move(Index)
End Sub
'+----------------------------------------------------------+
'+      各選択項目のMouse Move 処理
'+----------------------------------------------------------+
Private Sub Mouse_Move(Index As Integer)
    On Error GoTo Mouse_Move_Err

    ' 前回選択項目クリア
    For Each mCisMenu In POP_Menu
        If mCisMenu.Index = Index Then
            mCisMenu.cBoderWidth = 3
            mCisMenu.BackColor = &HFFC0C0
        Else
            mCisMenu.cBoderWidth = 1
            mCisMenu.BackColor = &HC0C0C0
        End If
    Next mCisMenu
'
Mouse_Move_Next:
    On Error GoTo 0

    Exit Sub
Mouse_Move_Err:
    GoTo Mouse_Move_Next
End Sub
'+------------------------------------------+
'+      ポップアップ処理（行単位処理）
'+------------------------------------------+
Private Sub Pop_Menu_Click(Index As Integer)

'挿入
    If Index = 0 Then
        If Not TorcdInsert_Proc Then Exit Sub
    End If
'ｺﾋﾟｰ
    If Index = 1 Then
        CopyTorcd = RTrim(vsGrid.TextMatrix(CopyRow, (CopyCol - 1) * 3 + 1))
        CopyUkeir = RTrim(vsGrid.TextMatrix(CopyRow, (CopyCol - 1) * 3 + 2))
        CopyTornm = RTrim(vsGrid.TextMatrix(CopyRow, (CopyCol - 1) * 3 + 3))
    End If
'貼り付け
    If Index = 2 Then
        Call Paste_Proc
    End If
'削除
    If Index = 3 Then
        If Not TorcdDelete_Proc Then Exit Sub
    End If
'クリア
    If Index = 4 Then
        vsGrid.TextMatrix(CopyRow, (CopyCol - 1) * 3 + 1) = ""
        vsGrid.TextMatrix(CopyRow, (CopyCol - 1) * 3 + 2) = ""
        vsGrid.TextMatrix(CopyRow, (CopyCol - 1) * 3 + 3) = ""
    End If
'戻る
    Call VsGridToText(vsGrid.Row)
    If Not iB1_Torcd(3).Visible Then Call GridTextEnabled(True)
    If ActObj.Visible Then ActObj.SetFocus
    PopupSMenu.Visible = False
End Sub
'*------------------------------------------------------*
'*    行　貼り付け                          　          *
'*------------------------------------------------------*
Private Function Paste_Proc()
    Paste_Proc = False

    With vsGrid
        .Redraw = False
        .TextMatrix(CopyRow, (CopyCol - 1) * 3 + 1) = RTrim(CopyTorcd)
        .TextMatrix(CopyRow, (CopyCol - 1) * 3 + 2) = RTrim(CopyUkeir)
        .TextMatrix(CopyRow, (CopyCol - 1) * 3 + 3) = RTrim(CopyTornm)
        .Redraw = True: .Col = 1
    End With

    Paste_Proc = True
End Function
'*------------------------------------------------------*
'*    取引先　削除　（取引先・受入を詰める）   　          *
'*------------------------------------------------------*
Private Function TorcdDelete_Proc()
    TorcdDelete_Proc = False
    Dim MotoRow     As Integer
    Dim MotoCol     As Integer
    Dim SakiRow     As Integer
    Dim SakiCol     As Integer
'行を詰める
    With vsGrid
        .Redraw = False
        If CopyRow > 5 Then
            gCnt1 = (CopyCol - 1) * 5 + CopyRow - 6 + 30
        Else
            gCnt1 = (CopyCol - 1) * 5 + CopyRow
        End If
        gCnt2 = gCnt1 + 1

        For gInt = gCnt2 To 60
            '移動先
            Call RowCol_Set(CInt(gCnt1))
            SakiRow = mSaveRow
            SakiCol = mSaveCol
            '移動元
            Call RowCol_Set(CInt(gCnt2))
            MotoRow = mSaveRow
            MotoCol = mSaveCol
            
            .TextMatrix(SakiRow, (SakiCol - 1) * 3 + 1) = RTrim(.TextMatrix(MotoRow, (MotoCol - 1) * 3 + 1))
            .TextMatrix(SakiRow, (SakiCol - 1) * 3 + 2) = RTrim(.TextMatrix(MotoRow, (MotoCol - 1) * 3 + 2))
            .TextMatrix(SakiRow, (SakiCol - 1) * 3 + 3) = RTrim(.TextMatrix(MotoRow, (MotoCol - 1) * 3 + 3))
            
            gCnt1 = gCnt1 + 1: gCnt2 = gCnt2 + 1
        Next
        
        .Redraw = True: .Col = 1
    End With
    SakiRow = 11: SakiCol = 6
    With vsGrid
        .TextMatrix(SakiRow, (SakiCol - 1) * 3 + 1) = ""
        .TextMatrix(SakiRow, (SakiCol - 1) * 3 + 2) = ""
        .TextMatrix(SakiRow, (SakiCol - 1) * 3 + 3) = ""
    End With

    TorcdDelete_Proc = True
End Function
'*------------------------------------------------------*
'*    取引先　挿入　（取引先・受入を後ろに移動）          *
'*------------------------------------------------------*
Private Function TorcdInsert_Proc()
    TorcdInsert_Proc = False
    Dim MotoRow     As Integer
    Dim MotoCol     As Integer
    Dim SakiRow     As Integer
    Dim SakiCol     As Integer
'行を詰める
    With vsGrid
        .Redraw = False
        If CopyRow > 5 Then
            gCnt1 = (CopyCol - 1) * 5 + CopyRow - 6 + 30
        Else
            gCnt1 = (CopyCol - 1) * 5 + CopyRow
        End If

        For gInt = 59 To gCnt1 Step -1
            '移動先
            Call RowCol_Set(gInt + 1)
            SakiRow = mSaveRow
            SakiCol = mSaveCol
            '移動元
            Call RowCol_Set(gInt)
            MotoRow = mSaveRow
            MotoCol = mSaveCol
            
            .TextMatrix(SakiRow, (SakiCol - 1) * 3 + 1) = RTrim(.TextMatrix(MotoRow, (MotoCol - 1) * 3 + 1))
            .TextMatrix(SakiRow, (SakiCol - 1) * 3 + 2) = RTrim(.TextMatrix(MotoRow, (MotoCol - 1) * 3 + 2))
            .TextMatrix(SakiRow, (SakiCol - 1) * 3 + 3) = RTrim(.TextMatrix(MotoRow, (MotoCol - 1) * 3 + 3))
            
        Next
        
        .Redraw = True: .Col = 1
    End With
    Call RowCol_Set(CInt(gCnt1))
    SakiRow = mSaveRow
    SakiCol = mSaveCol
    With vsGrid
        .TextMatrix(SakiRow, (SakiCol - 1) * 3 + 1) = ""
        .TextMatrix(SakiRow, (SakiCol - 1) * 3 + 2) = ""
        .TextMatrix(SakiRow, (SakiCol - 1) * 3 + 3) = ""
    End With

    TorcdInsert_Proc = True
End Function
'*-------------------------------------------------------------------------------*
'*    ポケット№よりRow値,Col値をセットする   戻り値 mSaveRow=gCnt1,mSaveCol=gCnt2 *
'*-------------------------------------------------------------------------------*
Private Function RowCol_Set(PoketNo As Integer)
    RowCol_Set = False
        
    mSaveRow = 0: mSaveCol = 0
    If PoketNo < 1 Or PoketNo > 60 Then Exit Function
    'Row
    If PoketNo Mod 5 = 0 Then
        mSaveRow = 5
    Else
        mSaveRow = PoketNo Mod 5
    End If
    If PoketNo > 30 Then
        mSaveRow = mSaveRow + 6
    End If
    'Grid Col
    mSaveCol = CisFun.HasuProc(PoketNo / 5, 切捨て)
    If PoketNo > 30 Then
        mSaveCol = mSaveCol - 6
    End If
    If PoketNo Mod 5 <> 0 Then
        mSaveCol = mSaveCol + 1
    End If

    RowCol_Set = True
End Function
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit2()
    With CisVsGrid2
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "PNO")
        Call .FixedSet(中中, 左中, あり, 17, 1, "エラー内容1")
        Call .FixedSet(中中, 右中, あり, 3, 2, "PN2")
        Call .FixedSet(中中, 左中, あり, 15, 3, "エラー内容2")
        Call .FixedSet(中中, 右中, あり, 3, 4, "PN3")
        Call .FixedSet(中中, 左中, あり, 15, 5, "エラー内容3")
        Call .FixedSet(中中, 右中, あり, 3, 6, "PN4")
        Call .FixedSet(中中, 左中, あり, 15, 7, "エラー内容4")
        Call .FixedSet(中中, 右中, あり, 3, 8, "PN5")
        Call .FixedSet(中中, 左中, あり, 15, 9, "エラー内容5")
        Call .FixedSet(中中, 右中, あり, 3, 10, "PN6")
        Call .FixedSet(中中, 左中, あり, 15, 11, "エラー内容6")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11")
        .InitDisp
    End With
    
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet2()
    Dim ID          As Integer
    Dim IDCol       As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Long
'
    gSL_Select = "SELECT 数字1,値名称 FROM 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = 'ポケットグループ発注'"
    gSL_Select = gSL_Select & "   AND 区分タイプ = ''"
    gSL_Select = gSL_Select & " UNION "
    gSL_Select = gSL_Select & " SELECT 数字3 数字1,値名称 FROM 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = 'かんばんエラー設定発注'"
    gSL_Select = gSL_Select & "     AND 区分タイプ = ''"
    gSL_Select = gSL_Select & "     AND 数字2 = 0"
    Call SYMRead(gSL_Select, 0)
' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid2
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid2.FixedGet(CisFun.Val2(ID))
        Next ID
        For ID = 1 To 5
            For IDCol = 0 To 10 Step 2
                .Row = ID: .Col = IDCol
                .CellBackColor = mMColor
            Next IDCol
        Next
    End With
    ID = 0: IDCol = 0
    Do Until Not SYM_RDSTS
        ID = ID + 1
        If ID > 5 Then
            ID = 1
            IDCol = IDCol + 2
        End If
        vsGrid2.TextMatrix(ID, IDCol) = SYM.数字1
        vsGrid2.TextMatrix(ID, IDCol + 1) = RTrim(SYM.値名称)
        mEPocket(SYM.数字1) = 1
'''        B1_ErrName(SYM.数字1).Visible = True
        
        Call RowCol_Set(CInt(SYM.数字1))
        vsGrid.Row = mSaveRow: vsGrid.Col = (mSaveCol - 1) * 3 + 1: vsGrid.CellBackColor = mEColor
        vsGrid.Row = mSaveRow: vsGrid.Col = (mSaveCol - 1) * 3 + 2: vsGrid.CellBackColor = mEColor
        vsGrid.Row = mSaveRow: vsGrid.Col = (mSaveCol - 1) * 3 + 3: vsGrid.CellBackColor = mEColor
        
        
        Call SYMReadNext
    Loop

    Call SYMClose
    With vsGrid2
         .Row = 1: .Col = 1
         .Redraw = True
    End With
End Sub
'+----------------------------------------------------------+
'+   グリッド初期表示
'+----------------------------------------------------------+
Private Sub GridInitDisp()
''''かんばん読取エラーポケット表示
'''    gSL_Select = " SELECT * FROM 名称マスタ"
'''    gSL_Select = gSL_Select & " WHERE 区分名称 = 'かんばんエラーポケット'"
'''    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
'''    gSL_Select = gSL_Select & " ORDER BY 値"
'''    If SYMRead(gSL_Select, 0) Then
'''        gInt = 0
'''        Do Until Not SYM_RDSTS
'''            gInt = gInt + 1
'''            B1_ErrName(gInt) = RTrim(SYM.値名称)
'''            Call SYMReadNext
'''        Loop
'''    End If
'''    Call SYMClose
    
    Erase mEPocket
'
'かんばん読取エラーポケット表示
    gSL_Select = " SELECT * FROM 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = 'ポケット最大数'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   And 値 = '1'"
    If SYMRead(gSL_Select, 1) Then
        mMColor = SYM.数字2
        mPocketMax = SYM.数字1
    End If
'
    Call GridColor
    Call GridSet2
    
End Sub
Private Sub GridColor()
    If mPocketMax < 60 Then
        With vsGrid
            .Redraw = False
            For gInt = mPocketMax + 1 To 60
                Call RowCol_Set(gInt)
                vsGrid.Row = mSaveRow: vsGrid.Col = (mSaveCol - 1) * 3 + 1: vsGrid.CellBackColor = mMColor
                vsGrid.Row = mSaveRow: vsGrid.Col = (mSaveCol - 1) * 3 + 2: vsGrid.CellBackColor = mMColor
                vsGrid.Row = mSaveRow: vsGrid.Col = (mSaveCol - 1) * 3 + 3: vsGrid.CellBackColor = mMColor
                mEPocket(gInt) = 2
            Next gInt
            .Redraw = True
        End With
    End If

End Sub


