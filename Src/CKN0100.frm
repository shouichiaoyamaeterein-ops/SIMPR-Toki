VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CKN0100 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "発注サイクルマスタ　メンテ"
   ClientHeight    =   10830
   ClientLeft      =   60
   ClientTop       =   1485
   ClientWidth     =   13860
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
   ScaleHeight     =   10830
   ScaleWidth      =   13860
   WindowState     =   2  '最大化
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'なし
      Height          =   345
      Index           =   1
      Left            =   8025
      ScaleHeight     =   345
      ScaleWidth      =   3675
      TabIndex        =   23
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
      Height          =   945
      Left            =   2070
      ScaleHeight     =   885
      ScaleWidth      =   5955
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1005
      Width           =   6015
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   780
         Left            =   5265
         Top             =   60
         Width           =   660
         _ExtentX        =   1164
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
            Left            =   135
            TabIndex        =   4
            Top             =   330
            Width           =   375
            _ExtentX        =   661
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
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   780
         Left            =   60
         Top             =   60
         Width           =   3660
         _ExtentX        =   6456
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
            Left            =   90
            TabIndex        =   0
            Top             =   330
            Width           =   1005
            _ExtentX        =   1773
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
            Left            =   1080
            Top             =   330
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   635
            BackColor       =   16777152
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
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D15 
         Height          =   780
         Index           =   0
         Left            =   3720
         Top             =   60
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1376
         ForeColor       =   16711680
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D CIS3D31 
            Height          =   195
            Left            =   390
            Top             =   330
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
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
         End
         Begin CisText_V60.CisText iH1_Cycle2 
            Height          =   360
            Left            =   525
            TabIndex        =   2
            Top             =   330
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
         Begin CisText_V60.CisText iH1_Cycle3 
            Height          =   360
            Left            =   1050
            TabIndex        =   3
            Top             =   330
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
         Begin CisText_V60.CisText iH1_Cycle1 
            Height          =   360
            Left            =   105
            TabIndex        =   1
            Top             =   330
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   635
            cFormat         =   "0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            MaxLength       =   1
            cDataType       =   1
            cDataReplace    =   1
            cFaZero         =   0
            cFdAutoFormat   =   1
            cGFormat        =   "#"
            cILength        =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D32 
            Height          =   195
            Left            =   915
            Top             =   330
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
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
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
      Height          =   1305
      Left            =   9810
      ScaleHeight     =   1245
      ScaleWidth      =   4275
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1020
      Width           =   4335
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   2565
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
         Left            =   3405
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
         Left            =   3405
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
         Left            =   1725
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
         Left            =   885
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
         Left            =   45
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
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   585
         Left            =   45
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   1725
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
      Begin CisBtn_60.CisBtn PB_PRT 
         Height          =   585
         Left            =   2565
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
      TabIndex        =   20
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
      Height          =   8310
      Left            =   2085
      ScaleHeight     =   8250
      ScaleWidth      =   7245
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2340
      Width           =   7305
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   8115
         Index           =   1
         Left            =   90
         Top             =   75
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   14314
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
         Begin vsFlexLib.vsFlexArray vsGrid 
            Height          =   7470
            Left            =   60
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   570
            Width           =   6945
            _Version        =   196608
            _ExtentX        =   12250
            _ExtentY        =   13176
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
            FormatString    =   "入荷便|昼夜勤　|振出時間|加算日|発行時間|回収時間|入荷時間||"
            Rows            =   17
            Cols            =   9
            BackColor       =   12648384
            BackColorBkg    =   12632256
            FocusRect       =   2
            HighLight       =   2
            ScrollBars      =   0
            RowHeightMin    =   430
            SelectionMode   =   1
            Begin Cis3D_v60.CIS3D Input_Row 
               Height          =   510
               Left            =   810
               Top             =   435
               Width           =   6075
               _ExtentX        =   10716
               _ExtentY        =   900
               BackColor       =   0
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
               cAlingnment     =   6
               cBoderStyle     =   1
               Begin Cis3D_v60.CIS3D Back_HY 
                  Height          =   420
                  Left            =   0
                  Top             =   0
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   741
                  BackColor       =   14737632
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
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   6
                  Begin VB.CheckBox ChkB1_HYKb 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "昼勤"
                     Height          =   255
                     Left            =   105
                     TabIndex        =   8
                     Top             =   90
                     Width           =   855
                  End
               End
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   420
                  Index           =   1
                  Left            =   1050
                  Top             =   0
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   16711680
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
                  cBoderColor1    =   -2147483643
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   7
                  Begin CisText_V60.CisText iB1_FHh 
                     Height          =   360
                     Left            =   90
                     TabIndex        =   9
                     Top             =   30
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
                     cGFormat        =   "00"
                     cILength        =   2
                  End
                  Begin CisText_V60.CisText iB1_FMm 
                     Height          =   360
                     Left            =   570
                     TabIndex        =   10
                     Top             =   30
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
                     cGFormat        =   "00"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   420
                  Index           =   2
                  Left            =   2910
                  Top             =   0
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   16711680
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
                  cBoderColor1    =   -2147483643
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   7
                  Begin CisText_V60.CisText iB1_HHh 
                     Height          =   360
                     Left            =   90
                     TabIndex        =   12
                     Top             =   30
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
                     cGFormat        =   "00"
                     cILength        =   2
                  End
                  Begin CisText_V60.CisText iB1_HMm 
                     Height          =   360
                     Left            =   570
                     TabIndex        =   13
                     Top             =   30
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
                     cGFormat        =   "00"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   420
                  Index           =   3
                  Left            =   3960
                  Top             =   0
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   16711680
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
                  cBoderColor1    =   -2147483643
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   7
                  Begin CisText_V60.CisText iB1_KHh 
                     Height          =   360
                     Left            =   90
                     TabIndex        =   14
                     Top             =   30
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
                     cGFormat        =   "00"
                     cILength        =   2
                  End
                  Begin CisText_V60.CisText iB1_KMm 
                     Height          =   360
                     Left            =   570
                     TabIndex        =   15
                     Top             =   30
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
                     cGFormat        =   "00"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   420
                  Index           =   4
                  Left            =   5010
                  Top             =   0
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   16711680
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
                  cBoderColor1    =   -2147483643
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   7
                  Begin CisText_V60.CisText iB1_NHh 
                     Height          =   360
                     Left            =   90
                     TabIndex        =   16
                     Top             =   30
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
                     cGFormat        =   "00"
                     cILength        =   2
                  End
                  Begin CisText_V60.CisText iB1_NMm 
                     Height          =   360
                     Left            =   570
                     TabIndex        =   17
                     Top             =   30
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
                     cGFormat        =   "00"
                     cILength        =   2
                  End
               End
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   420
                  Index           =   5
                  Left            =   2100
                  Top             =   0
                  Width           =   810
                  _ExtentX        =   1429
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   "　  日"
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
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   7
                  Begin CisText_V60.CisText iB1_KDd 
                     Height          =   360
                     Left            =   135
                     TabIndex        =   11
                     Top             =   30
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
            End
         End
         Begin CisText_V60.CisText B_TabDummy 
            Height          =   165
            Left            =   5340
            TabIndex        =   18
            Top             =   855
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   291
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   540
            Left            =   15
            Top             =   15
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   953
            ForeColor       =   16711680
            Caption         =   "最終発注日"
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
            cPositionX      =   30
            Begin CisYMD_With_Btn.CisYMDwB iB1_SEHYmd 
               Height          =   330
               Left            =   1275
               TabIndex        =   5
               Top             =   120
               Width           =   1695
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
            Begin CisText_V60.CisText iB1_SEHBin 
               Height          =   360
               Left            =   2985
               TabIndex        =   6
               Top             =   120
               Width           =   345
               _ExtentX        =   609
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
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   345
               Left            =   3300
               Top             =   150
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   609
               ForeColor       =   0
               Caption         =   "便"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   1
               cBoderStyle     =   2
               cPositionX      =   60
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   540
            Left            =   3825
            Top             =   15
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   953
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cPositionX      =   60
         End
         Begin CisText_V60.CisText B_TabDummy2 
            Height          =   165
            Left            =   375
            TabIndex        =   7
            Top             =   585
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   291
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      Caption         =   "【 発注サイクルマスタ　メンテ  】"
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
Attribute VB_Name = "CKN0100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   発注サイクルマスタメンテナンス
'**       フォームID    :   CKN0100
'**       処理概要      :
'**
'**       作  成  日    :   2003/12/02  By CIS SIMPR-A
'**       変  更  日    :   2004/05/11  By CIS - 1.印刷処理の追加：発注サイクル一覧表
'**       変  更  日    :   2004/05/14  By CIS 最終発注日追加
'**       変  更  日    :   2004/12/21  By CIS 振出時間"0:00"の不具合対応
'**       変  更  日    :   2005/02/21  By CIS 最終発注日の妥当性チェックを追加
'**       変  更  日    :   2005/03/04  By CIS 発注テーブルによる納入日算出方法追加による変更
'**                                            ｼｽﾃﾑｵﾌﾟｼｮﾝの納入日算出方法が発注テーブルの時、
'**                                            ｻｲｸﾙ(日)が１のサイクルのメンテは不可
'**                                            発注サイクルマスタに項目追加（最終発注便・発注処理回数・最終便）
'**       変  更  日    :   2005/03/07  By CIS ｶｰｿﾙが明細行にある時、入力ﾎﾞﾀﾝをｸﾘｯｸ時『実行時ｴﾗｰ '5'』になる対応
'**       変  更  日    :   2005/04/20  By CIS 発注テーブルによる納入日算出の時は最終発注日・便のみ変更可能
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)
    Dim SysOption       As Byte         '納入日便算出方法

    Dim ActObj          As Object
    Dim ErrObj          As Object
    Dim ActRow          As Integer
    Dim BinMax          As Byte         '設定最大便（８便）
    Dim mCycle          As String       'サイクル退避
    
    Dim mFTime          As String       '振出時間退避
    Dim mHTime          As String       '発行時間退避
    Dim mKTime          As String       '回収時間退避
    Dim mNTime          As String       '入荷時間退避
    Dim mBinTime()      As String       '便設定時間
    Dim mSTTime()       As String       '開始時間
    Dim mEDTime()       As String       '終了時間
    Dim mFurBin()       As Byte         '振出便
    Dim mEdaNo()        As Byte         '行識別
    Dim mHCDay()        As Byte         '発注加算日
    Dim mHCBin()        As Byte         '発注便
    Dim mFurDay()       As Integer      '振出加算日
    Dim mNYDay()        As Byte         '納入加算日
    Dim mHYKbn()        As Byte         '昼夜区分
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
    
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = vsGrid
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 430
        .PatanMax = 1
        .SelectCol = 19
        .InitGet = False
        .Init
    End With
    Call GridInit
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = "MNT"
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
'
    BinMax = 16
    ReDim mBinTime(BinMax)
    ReDim mSTTime(BinMax + 1): ReDim mEDTime(BinMax + 1): ReDim mFurBin(BinMax + 1)
    ReDim mEdaNo(BinMax + 1): ReDim mHCDay(BinMax + 1): ReDim mHCBin(BinMax + 1)
    ReDim mFurDay(BinMax + 1): ReDim mNYDay(BinMax + 1): ReDim mHYKbn(BinMax + 1)
    
' 納入日・便算出方法(システム設定値)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 14
    If Not IRNRead("", 1) Then
        IRN.oyian = "0"
    End If
    SysOption = CisFun.Val2(IRN.oyian)
    
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
           Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click    '【新規】
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    '【修正】
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyF8:       If PB_PRT.Visible Then Call PB_PRT_Click    '【印刷】
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      '【入力】
'▽▽:::::::::::::::::::::::::::::::::::::::::::::::::::::::<< 2005/03/07 >> Update Start ::::::::::
'                                 Dummy.Enabled = True: Dummy.SetFocus: DoEvents
'                                 Call PB_ENT_Click
'                                 Dummy.Enabled = False
                                  Call PB_ENT_Click
'△△:::::::::::::::::::::::::::::::::::::::::::::::::::::::<< 2005/03/07 >> Update End   ::::::::::
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
'▽▽:::::::::::::::::::::::::::::::::::::::::::::::::::::::<< 2005/03/07 >> Update Start ::::::::::
'   Call ReturnPress
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call ReturnPress
    Dummy.Enabled = False
'△△:::::::::::::::::::::::::::::::::::::::::::::::::::::::<< 2005/03/07 >> Update End   ::::::::::
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
   Call CisFrm.MousePT(11)
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
   
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKB)
      iH1_Torcd.SetFocus
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Dummy.Enabled = False
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
    iH1_Torcd.SetFocus
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If SyoriKB <> "MNT" Then
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
    End If
    iH1_Torcd.SetFocus
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
    iH1_Torcd.SetFocus
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = " SELECT *"
    gSL_Select = gSL_Select & " FROM 発注サイクルマスタ "
    gSL_Select = gSL_Select & " WHERE 仕入先+サイクル+受入 < '" & RTrim(iH1_Torcd)
    gSL_Select = gSL_Select & RTrim(iH1_Cycle1) & RTrim(iH1_Cycle2) & RTrim(iH1_Cycle3) & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & " ORDER BY 仕入先 Desc,サイクル desc,受入 desc"
    If Not HCMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     先頭レコードです            "
            .MB_Title = "【後退】"
            .MB_Button = OK
            .MBOX
        End With
        GoTo PB_Back_Ed
    End If
    gSL_Select = " SELECT * "
    gSL_Select = gSL_Select & " FROM 発注サイクルマスタ "
    gSL_Select = gSL_Select & " WHERE 仕入先 = '" & RTrim(HCM.仕入先) & "'"
    gSL_Select = gSL_Select & "   And サイクル = '" & RTrim(HCM.サイクル) & "'"
    gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(HCM.受入) & "'"
    gSL_Select = gSL_Select & "   And 行識別 = 0"
    gSL_Select = gSL_Select & " ORDER BY 振出便"
    Call DBInput("Back")
PB_Back_Ed:
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Torcd.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = " SELECT *"
    gSL_Select = gSL_Select & " FROM 発注サイクルマスタ "
    gSL_Select = gSL_Select & " WHERE 仕入先+サイクル+受入 > '" & RTrim(iH1_Torcd)
    gSL_Select = gSL_Select & RTrim(iH1_Cycle1) & RTrim(iH1_Cycle2) & RTrim(iH1_Cycle3) & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & " ORDER BY 仕入先,サイクル,受入 "
    If Not HCMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     最終レコードです            "
            .MB_Title = "【前進】"
            .MB_Button = OK
            .MBOX
        End With
        GoTo PB_Next_Ed
    End If
    gSL_Select = " SELECT * "
    gSL_Select = gSL_Select & " FROM 発注サイクルマスタ "
    gSL_Select = gSL_Select & " WHERE 仕入先 = '" & RTrim(HCM.仕入先) & "'"
    gSL_Select = gSL_Select & "   And サイクル = '" & RTrim(HCM.サイクル) & "'"
    gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(HCM.受入) & "'"
    gSL_Select = gSL_Select & "   And 行識別 = 0"
    gSL_Select = gSL_Select & " ORDER BY 振出便"
    Call DBInput("Next")
PB_Next_Ed:
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Torcd.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    印　刷  キ ー(F8)       +
'+----------------------------+
Private Sub PB_PRT_Click()
'   リスト出力先確認
    CKN0110.Show vbModal
    Unload CKN0110
    Set CKN0110 = Nothing
'   印刷処理
    If CKN0110_PrintOut <> 9 Then
        Call CKN011nLPrint(CKN0110_PrintKB, CKN0110_PrintOut)
    End If
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Cycle1" Then Call Look_Cycle: Exit Sub
    If PB_Look.Tag = "iH1_Cycle2" Then Call Look_Cycle: Exit Sub
    If PB_Look.Tag = "iH1_Cycle3" Then Call Look_Cycle: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_Ukeir: Exit Sub
    If PB_Look.Tag = "iB1_SEHYmd" Then Call iB1_SEHYmd.ShowCalender: Exit Sub
End Sub
'+----------------------------+
'+       取引先検索            +
'+----------------------------+
Private Sub Look_Torcd()

    RV_Left = 0:    RV_Top = 0
    RV_TorKb = 8: RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
       iH1_Torcd = RTrim(RV_TorcdK)
       Call iH1_Torcd_LostFocus
    End If

End Sub
'+----------------------------+
'+       受入検索            +
'+----------------------------+
Private Sub Look_Ukeir()

    RV_Left = 0: RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 8
    RV_Torcd = iH1_Torcd
    RV_Kojyo = ""
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
       iH1_Torcd = RTrim(RV_TorcdK)
       Call iH1_Torcd_LostFocus
       iH1_Ukeir = RTrim(RV_Ukeir)
    End If

End Sub
'+----------------------------+
'+       発注サイクル検索      +
'+----------------------------+
Private Sub Look_Cycle()

    RV_Left = 0:    RV_Top = 0
    RV_Call = "C"
    RV_Torcd = "": RV_TorcdK = "": RV_Ukeir = "": RV_Kojyo = ""
    
    CKNK0100.Show vbModal
    Unload CKNK0100
    Set CKNK0100 = Nothing

    If RV_Rtn Then
       iH1_Torcd = RTrim(RV_Torcd)
       Call iH1_Torcd_LostFocus
       iH1_Cycle1 = Left(RV_Kojyo, 1)
       iH1_Cycle2 = Mid(RV_Kojyo, 2, 2)
       iH1_Cycle3 = Mid(RV_Kojyo, 4, 2)
       iH1_Ukeir = RTrim(RV_Ukeir)
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
    If RTrim(iH1_Torcd) <> "" Then
        If TorNmGet(iH1_Torcd, 8) Then
            H1lb_Tornm = RTrim(TRM.略称)
        End If
    End If
End Sub
'+-----------------------------+
'+  iH1_Cycle1
'+-----------------------------+
Private Sub iH1_Cycle1_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Cycle1_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    iH1_Cycle1 = Format(CisFun.Val2(iH1_Cycle1), "0")
End Sub
'+-----------------------------+
'+  iH1_Cycle2
'+-----------------------------+
Private Sub iH1_Cycle2_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Cycle2_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    iH1_Cycle2 = Format(CisFun.Val2(iH1_Cycle2), "00")
End Sub
'+-----------------------------+
'+  iH1_Cycle3
'+-----------------------------+
Private Sub iH1_Cycle3_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Cycle3_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    iH1_Cycle3 = Format(CisFun.Val2(iH1_Cycle3), "00")
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
'+  iB1_SEHYmd
'+-----------------------------+
Private Sub iB1_SEHYmd_GotFocus()
    Back_HY.Tag = ""
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_SEHYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  ChkB1_HYKb
'+-----------------------------+
Private Sub ChkB1_HYKb_GotFocus()
    Set ActObj = ChkB1_HYKb
    Back_HY.Tag = "S"
End Sub
Private Sub ChkB1_HYKb_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub ChkB1_HYKb_LostFocus()
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 1) = ChkB1_HYKb.Caption
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 7) = ChkB1_HYKb.Value
End Sub
Private Sub ChkB1_HYKb_Click()
    If ChkB1_HYKb.Value Then
        ChkB1_HYKb.Caption = "夜勤"
    Else
        ChkB1_HYKb.Caption = "昼勤"
    End If
End Sub
'+-----------------------------+
'+  iB1_FHh
'+-----------------------------+
Private Sub iB1_FHh_GotFocus()
    Set ActObj = iB1_FHh
End Sub
Private Sub iB1_FHh_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_FHh_LostFocus()
    Call EditTime(iB1_FHh, iB1_FMm, gStr)
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 2) = gStr
End Sub
'+-----------------------------+
'+  iB1_FMm
'+-----------------------------+
Private Sub iB1_FMm_GotFocus()
    Set ActObj = iB1_FMm
End Sub
Private Sub iB1_FMm_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_FMm_LostFocus()
    Call EditTime(iB1_FHh, iB1_FMm, gStr)
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 2) = gStr
End Sub
'+-----------------------------+
'+  iB1_KDd
'+-----------------------------+
Private Sub iB1_KDd_GotFocus()
    Set ActObj = iB1_KDd
End Sub
Private Sub iB1_KDd_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_KDd_LostFocus()
    If CisFun.Val2(iB1_KDd) <> 0 Then
        vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 3) = iB1_KDd
    Else
        vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 3) = ""
    End If
End Sub
'+-----------------------------+
'+  iB1_HHh
'+-----------------------------+
Private Sub iB1_HHh_GotFocus()
    Set ActObj = iB1_HHh
End Sub
Private Sub iB1_HHh_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_HHh_LostFocus()
    Call EditTime(iB1_HHh, iB1_HMm, gStr)
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 4) = gStr
End Sub
'+-----------------------------+
'+  iB1_HMm
'+-----------------------------+
Private Sub iB1_HMm_GotFocus()
    Set ActObj = iB1_HMm
End Sub
Private Sub iB1_HMm_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_HMm_LostFocus()
    Call EditTime(iB1_HHh, iB1_HMm, gStr)
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 4) = gStr
End Sub
'+-----------------------------+
'+  iB1_KHh
'+-----------------------------+
Private Sub iB1_KHh_GotFocus()
    Set ActObj = iB1_KHh
End Sub
Private Sub iB1_KHh_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_KHh_LostFocus()
    Call EditTime(iB1_KHh, iB1_KMm, gStr)
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 5) = gStr
End Sub
'+-----------------------------+
'+  iB1_KMm
'+-----------------------------+
Private Sub iB1_KMm_GotFocus()
    Set ActObj = iB1_KMm
End Sub
Private Sub iB1_KMm_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_KMm_LostFocus()
    Call EditTime(iB1_KHh, iB1_KMm, gStr)
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 5) = gStr
End Sub
'+-----------------------------+
'+  iB1_NHh
'+-----------------------------+
Private Sub iB1_NHh_GotFocus()
    Set ActObj = iB1_NHh
End Sub
Private Sub iB1_NHh_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_NHh_LostFocus()
    Call EditTime(iB1_NHh, iB1_NMm, gStr)
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 6) = gStr
End Sub
'+-----------------------------+
'+  iB1_NMm
'+-----------------------------+
Private Sub iB1_NMm_GotFocus()
    Set ActObj = iB1_NMm
    Back_HY.Tag = "E"
End Sub
Private Sub iB1_NMm_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_NMm_LostFocus()
    Call EditTime(iB1_NHh, iB1_NMm, gStr)
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 6) = gStr
End Sub
Private Function EditTime(EditHh As String, EditMm As String, EditHHMM As String)
    EditHHMM = Space(5)
    
    If RTrim(EditHh) <> "" Then
        EditHh = Format(CisFun.Val2(EditHh), "00")
    End If
    If RTrim(EditMm) <> "" Then
        EditMm = Format(CisFun.Val2(EditMm), "00")
    End If
    
    If RTrim(EditHh) <> "" And RTrim(EditMm) = "" Then
        EditMm = "00"
    End If
    If RTrim(EditHh) = "" And RTrim(EditMm) <> "" Then
        EditMm = ""
    End If
    
    If RTrim(EditHh) <> "" Then
        EditHHMM = Format(CisFun.Val2(EditHh), "00")
    Else
        EditHHMM = "  "
    End If
    If RTrim(EditMm) <> "" Then
        EditHHMM = EditHHMM & ":" & Format(CisFun.Val2(EditMm), "00")
    Else
        EditHHMM = EditHHMM & "   "
    End If
End Function
'+-----------------------------+
'+  B_TabDummy2
'+-----------------------------+
Private Sub B_TabDummy2_GotFocus()
    If Back_HY.Tag = "" Then
        Call TextToVsGrid(Val(Input_Row.Tag))
        Call Grid_RowChange(vbKeyUp)
        ChkB1_HYKb.SetFocus
        Exit Sub
    Else
        If Val(Input_Row.Tag) = 1 Then
           iB1_SEHYmd.SetFocus
        Else
           Call TextToVsGrid(Val(Input_Row.Tag))
           Call Grid_RowChange(vbKeyUp)
           iB1_NMm.SetFocus
        End If
    End If
End Sub
'+-----------------------------+
'+  B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
    If Back_HY.Tag = "S" Then
        Call TextToVsGrid(Val(Input_Row.Tag))
        Call Grid_RowChange(vbKeyUp)
    End If
    If Back_HY.Tag = "E" Then
        Call TextToVsGrid(Val(Input_Row.Tag))
        Call Grid_RowChange(vbKeyDown)
    End If
    'SendKeys "{Tab}"
    ChkB1_HYKb.SetFocus
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
     Case "LOK"
          SyoriNM = "照会"
          K_Sykbnm.ForeColor = &HFF00FF   '(ピンク)
          PB_ADD.Visible = False
          PB_MNT.Visible = False
          PB_DEL.Visible = False
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
    PB_PRT.Visible = H1Mode

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
            If .Name Like "Chk" & wbc_Enm & "_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
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
        If Not Head1Chk Then GoTo ReturnPress_Ed
' << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call HeadBodyClear("B")
            ProcHB = "H1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
        Else
            ProcHB = "B1"
            Call DispChange(ProcHB)
            vsGrid.Row = 1
            Call VsGridToText(vsGrid.Row)
        End If
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
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
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
'仕入先
    If Trim(iH1_Torcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    仕入先を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    End If
    If Not TorNmGet(iH1_Torcd, 8) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    仕入先　未登録！（取引先マスタ）            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    End If
'サイクル
    If Val(iH1_Cycle1) = 0 Or Val(iH1_Cycle2) = 0 Or Val(iH1_Cycle3) = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    サイクルを再入力して下さい！            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Cycle1.SetFocus
       Exit Function
    End If
    If Val(iH1_Cycle2) > 16 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    サイクル[回]は １６ までです！            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Cycle2.SetFocus
       Exit Function
    End If
    If Val(iH1_Cycle1) <> 1 And Val(iH1_Cycle2) > 1 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    サイクル[日] >１ の時はサイクル[回]は １限定です！            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Cycle2.SetFocus
       Exit Function
    End If
    mCycle = Format(Val(iH1_Cycle1), "0") & Format(Val(iH1_Cycle2), "00") & Format(Val(iH1_Cycle3), "00")

'( 納入日算出方法による制限 )              ---[ 2005/03/04 Start ]
'''    If SysOption <> 0 And Val(iH1_Cycle1) = 1 And SyoriKB <> "DEL" Then
'''        With CisFun
'''            .MB_Lines = 4
'''            .MB_MSG(2) = "    納入日･便の算出方法は [発注テーブル] の設定になっています！            "
'''            .MB_MSG(4) = "    サイクル(日)が１のサイクルは発注テーブルメンテで設定を行って下さい。            "
'''            .MB_Title = ""
'''            .MB_Button = Error
'''            .MBOX
'''        End With
'''       iH1_Cycle1.SetFocus
'''       Exit Function
'''    End If
'-------------------------------------------[ 2005/03/04 End  ]
'受入
    If RTrim(iH1_Ukeir) = "" Then
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
    If RTrim(iH1_Ukeir) <> "" Then
        TUM.取引先CD = RTrim(iH1_Torcd)
        TUM.受入 = RTrim(iH1_Ukeir)
        TUM.取引先区分 = 1
        If Not TUMRead("", 1) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    受入　未登録！（取引先マスタ）            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_Torcd.SetFocus
           Exit Function
        End If
    End If
' マスタ入力 SELECT文 ｾｯﾄ
    gSL_Select = " SELECT * "
    gSL_Select = gSL_Select & " FROM 発注サイクルマスタ "
    gSL_Select = gSL_Select & " WHERE 仕入先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   And サイクル = '" & RTrim(mCycle) & "'"
    gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   And 行識別 = 0"
    gSL_Select = gSL_Select & " ORDER BY 振出便"
    If DBInput("INV") Then
        If SyoriKB = "ADD" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    発注サイクルマスタ登録済！           "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Torcd.SetFocus
            Exit Function
        End If
    Else
        If SyoriKB <> "ADD" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    発注サイクルマスタ未登録！           "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Torcd.SetFocus
            Exit Function
        Else
            vsGrid.Redraw = False
            For gInt = 1 To CisFun.Val2(iH1_Cycle2)
                If gInt + 1 > vsGrid.Rows Then vsGrid.Rows = gInt + 1
                vsGrid.TextMatrix(gInt, 0) = gInt & "便"
            Next
            vsGrid.Redraw = True: vsGrid.Row = 1
        End If
    End If
       
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
    Dim YMD     As Date
    Dim wCol    As Byte
    Dim wBIdx   As Byte
    Dim w24h    As Byte
    Dim wHYKb   As Byte
    Dim wHYCnt  As Byte
    
    Call TextToVsGrid(Val(Input_Row.Tag))
    Call GridTextEnabled(False)

'<< 振出時間 >>チェック
    mFTime = Space(64): wCol = 1: w24h = 0: wBIdx = 0
    Erase mBinTime: ReDim mBinTime(BinMax): ReDim mHYKbn(BinMax + 1)
    
    wHYKb = CisFun.Val2(vsGrid.TextMatrix(1, 7)): wHYCnt = 0
'<< 振出時間 >>チェック
    For gCnt1 = 1 To CisFun.Val2(iH1_Cycle2)
        If wHYKb <> CisFun.Val2(vsGrid.TextMatrix(gCnt1, 7)) Then
            wHYCnt = wHYCnt + 1
            If wHYCnt > 2 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    昼夜勤区分　設定異常！            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = ChkB1_HYKb
                GoTo Body1Chk_Err
            End If
            wHYKb = CisFun.Val2(vsGrid.TextMatrix(gCnt1, 7))
        End If
        
        mBinTime(gCnt1) = ""
        iB1_FHh = Left(vsGrid.TextMatrix(gCnt1, 2), 2)
        iB1_FMm = Mid(vsGrid.TextMatrix(gCnt1, 2), 4, 2)
        If CisFun.Val2(iB1_FHh) <> 0 Then
            If CisFun.Val2(iB1_FHh) >= 24 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    振出時間･時 再入力            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iB1_FHh
                GoTo Body1Chk_Err
            End If
        End If
        If CisFun.Val2(iB1_FMm) <> 0 Then
            If CisFun.Val2(iB1_FMm) >= 60 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    振出時間･分 再入力            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iB1_FMm
                GoTo Body1Chk_Err
            End If
        End If
        If CisFun.Val2(iB1_FHh) = 0 Then
            iB1_FHh = "00"
        End If
        If CisFun.Val2(iB1_FMm) = 0 Then
            iB1_FMm.Text = "00"
        End If
        Mid(mFTime, wCol, 2) = Format(CisFun.Val2(iB1_FHh), "00")
        Mid(mFTime, wCol + 2, 2) = Format(CisFun.Val2(iB1_FMm), "00")
        vsGrid.TextMatrix(gCnt1, 2) = Format(CisFun.Val2(iB1_FHh), "00") & ":" & Format(CisFun.Val2(iB1_FMm), "00")
        mBinTime(gCnt1) = Mid(mFTime, wCol, 4)
        mHYKbn(gCnt1) = CisFun.Val2(vsGrid.TextMatrix(gCnt1, 7))
'振出時間設定　範囲重複チェック
        wBIdx = wBIdx + 1
        If wBIdx <> 1 Then
            If mBinTime(gCnt1) <= mBinTime(gCnt1 - 1) Then
                If w24h >= 1 Then
                    With CisFun
                        .MB_Lines = 4
                        .MB_MSG(2) = "    　時間設定エラー（時間重複）            "
                        .MB_Title = ""
                        .MB_Button = Error
                        .MBOX
                    End With
                    Set ErrObj = iB1_FHh
                    GoTo Body1Chk_Err
                End If
                w24h = w24h + 2
            End If
            For gCnt2 = 1 To wBIdx - 1
                If mBinTime(gCnt1) = mBinTime(gCnt2) Then
                    With CisFun
                        .MB_Lines = 4
                        .MB_MSG(2) = "    　時間設定エラー（同一時間有り）            "
                        .MB_Title = ""
                        .MB_Button = Error
                        .MBOX
                    End With
                    Set ErrObj = iB1_FHh
                    GoTo Body1Chk_Err
                End If
            Next
            If mBinTime(gCnt1) >= mBinTime(1) Then
                If w24h >= 2 Then
                    With CisFun
                        .MB_Lines = 4
                        .MB_MSG(2) = "      時間設定エラー（時間重複）            "
                        .MB_Title = ""
                        .MB_Button = Error
                        .MBOX
                    End With
                    Set ErrObj = iB1_FHh
                    GoTo Body1Chk_Err
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
        If mBinTime(1) > mBinTime(CisFun.Val2(iH1_Cycle2)) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "      １便が夜勤の時、1便振出時間　>　最終便振出時間            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            Set ErrObj = iB1_FHh
            gCnt1 = 1
            GoTo Body1Chk_Err
        End If
    End If
'<<振出便設定>>振出開始時間･終了時間ｾｯﾄ
    Call FuriBin_Set
'<< 発行時間 >>チェック
    mHTime = Space(64): wCol = 1
    For gCnt1 = 1 To CisFun.Val2(iH1_Cycle2)
        iB1_HHh = Left(vsGrid.TextMatrix(gCnt1, 4), 2)
        iB1_HMm = Mid(vsGrid.TextMatrix(gCnt1, 4), 4, 2)
        If CisFun.Val2(iB1_HHh) <> 0 Then
            If CisFun.Val2(Trim(iB1_HHh)) >= 24 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    発行時間･時 再入力            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iB1_HHh
                GoTo Body1Chk_Err
            End If
        End If
        If CisFun.Val2(iB1_HMm) <> 0 Then
            If CisFun.Val2(Trim(iB1_HMm)) >= 60 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    発行時間･分 再入力            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iB1_HMm
                GoTo Body1Chk_Err
            End If
        End If
        If CisFun.Val2(iB1_HHh) = 0 Then
            iB1_HHh = "00"
        End If
        If CisFun.Val2(iB1_HMm) = 0 Then
            iB1_HMm = "00"
        End If

        Mid(mHTime, wCol, 2) = Format(CisFun.Val2(iB1_HHh), "00")
        Mid(mHTime, wCol + 2, 2) = Format(CisFun.Val2(iB1_HMm), "00")
        vsGrid.TextMatrix(gCnt1, 4) = Format(CisFun.Val2(iB1_HHh), "00") & ":" & Format(CisFun.Val2(iB1_HMm), "00")
        wCol = wCol + 4
    Next gCnt1
'<< 回収時間 >>チェック
    mKTime = Space(64): wCol = 1
    For gCnt1 = 1 To CisFun.Val2(iH1_Cycle2)
        iB1_KHh = Left(vsGrid.TextMatrix(gCnt1, 5), 2)
        iB1_KMm = Mid(vsGrid.TextMatrix(gCnt1, 5), 4, 2)
        If CisFun.Val2(iB1_KHh) <> 0 Then
            If CisFun.Val2(Trim(iB1_KHh)) >= 24 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    回収時間･時 再入力            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iB1_KHh
                GoTo Body1Chk_Err
            End If
        End If
        If CisFun.Val2(iB1_KMm) <> 0 Then
            If CisFun.Val2(Trim(iB1_KMm)) >= 60 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    回収時間･分 再入力            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iB1_KMm
                GoTo Body1Chk_Err
            End If
        End If
        If CisFun.Val2(iB1_KHh) <> 0 And CisFun.Val2(iB1_KMm) = 0 Then
            iB1_KMm = "00"
        End If
        If iB1_KHh = "00" And CisFun.Val2(iB1_KMm) = 0 Then
            iB1_KMm = "00"
        End If
        If RTrim(iB1_KHh) = "" Then
            iB1_KMm = ""
        End If

        Mid(mKTime, wCol, 2) = Trim(iB1_KHh)
        Mid(mKTime, wCol + 2, 2) = Trim(iB1_KMm)
        If RTrim(Mid(mKTime, wCol, 4)) <> "" Then
            vsGrid.TextMatrix(gCnt1, 5) = Format(CisFun.Val2(iB1_KHh), "00") & ":" & Format(CisFun.Val2(iB1_KMm), "00")
        Else
            vsGrid.TextMatrix(gCnt1, 5) = ""
        End If
        wCol = wCol + 4
    Next gCnt1
'<< 入荷時間 >>チェック
    mNTime = Space(64): wCol = 1
    For gCnt1 = 1 To CisFun.Val2(iH1_Cycle2)
        iB1_NHh = Left(vsGrid.TextMatrix(gCnt1, 6), 2)
        iB1_NMm = Mid(vsGrid.TextMatrix(gCnt1, 6), 4, 2)
        If CisFun.Val2(iB1_NHh) <> 0 Then
            If CisFun.Val2(Trim(iB1_NHh)) >= 24 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    入荷時間･時 再入力            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iB1_NHh
                GoTo Body1Chk_Err
            End If
        End If
        If CisFun.Val2(iB1_NMm) <> 0 Then
            If CisFun.Val2(Trim(iB1_NMm)) >= 60 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    入荷時間･分 再入力            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                Set ErrObj = iB1_NMm
                GoTo Body1Chk_Err
            End If
        End If
        If CisFun.Val2(iB1_NHh) <> 0 And CisFun.Val2(iB1_NMm) = 0 Then
            iB1_NMm = "00"
        End If
        If iB1_NHh = "00" And CisFun.Val2(iB1_NMm) = 0 Then
            iB1_NMm = "00"
        End If
        If RTrim(iB1_NHh) = "" Then
            iB1_NMm = ""
        End If

        Mid(mNTime, wCol, 2) = Trim(iB1_NHh)
        Mid(mNTime, wCol + 2, 2) = Trim(iB1_NMm)
        If RTrim(Mid(mNTime, wCol, 4)) <> "" Then
            vsGrid.TextMatrix(gCnt1, 6) = Format(CisFun.Val2(iB1_NHh), "00") & ":" & Format(CisFun.Val2(iB1_NMm), "00")
        Else
            vsGrid.TextMatrix(gCnt1, 6) = ""
        End If
        wCol = wCol + 4
    Next gCnt1
    
    ' 最終発注日を求める
    If Trim(iB1_SEHYmd) = "" Then
       gSL_Select = "Select 開始発注日 "
       gSL_Select = gSL_Select & " FROM 取引先マスタ "
       gSL_Select = gSL_Select & " Where 取引先CD = '" & RTrim(iH1_Torcd) & "'"
       If Not TRMRead(gSL_Select, 1) Then
          Call ItemsClearTRM
       End If
       
       If CisFun.Val2(iH1_Cycle1) > 1 Then
          If Trim(TRM.開始発注日) = "" Then
             With CisFun
                  .MB_Lines = 4
                  .MB_MSG(2) = "    仕入先　開始発注日が入力されていません          "
                  .MB_MSG(4) = "        最終発注日を入力して下さい                  "
                  .MB_Title = ""
                  .MB_Button = Error
                  .MBOX
             End With
             Set ErrObj = iB1_SEHYmd
             gCnt1 = 1
             GoTo Body1Chk_Err
          End If
       End If
       If Not GetEndHDate Then
          Set ErrObj = iB1_SEHYmd
          gCnt1 = 1
          GoTo Body1Chk_Err
       End If
       
       If CisFun.Val2(iH1_Cycle1) > 1 Then
          If Trim(iB1_SEHYmd) = "" Then
             With CisFun
                  .MB_Lines = 4
                  .MB_MSG(2) = "    最終発注日を入力して下さい          "
                  .MB_Title = ""
                  .MB_Button = Error
                  .MBOX
             End With
             Set ErrObj = iB1_SEHYmd
             gCnt1 = 1
             GoTo Body1Chk_Err
          End If
       End If
'▽▽------------------------------------------------------<< 2005/02/21 >> Update Start
    Else
       If Not iB1_SEHYmd.cChkResult Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    最終発注日を正しく入力して下さい            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iB1_SEHYmd.SetFocus
          Exit Function
       End If
'△△------------------------------------------------------<< 2005/02/21 >> Update End
    End If
    
    Body1Chk = True
    Exit Function
Body1Chk_Err:
    vsGrid.Row = gCnt1
    Call VsGridToText(vsGrid.Row)
    Call GridTextEnabled(True)
    ErrObj.SetFocus
End Function
'+-----------------------------+
'+  最終発注日を求める
'+-----------------------------+
Private Function GetEndHDate() As Boolean
    Dim wHDate      As String
    Dim wDate       As String
    
    GetEndHDate = False
    
    wHDate = TRM.開始発注日
    If Trim(wHDate) = "" Then
       GetEndHDate = True
       Exit Function
    End If
    
    wDate = Format(Date, "YYYYMMDD")
    
    Do Until wDate <= wHDate
        With CisDB
        '   <><><> 最終発注日の獲得 <><><>
            .SQL = "日付加算"
            .StoadoCount = 4
            .ParaNo = 0: .ParaIO = Return用
            .ParaNo = 1: .ParaIO = Input用
            .ParaNo = 2: .ParaIO = Input用
            .ParaNo = 3: .ParaIO = OutPut用
            .ParaNo = 4: .ParaIO = OutPut用
        '
            .ParaNo = 1: .ParaValue = wHDate
            .ParaNo = 2: .ParaValue = CisFun.Val2(iH1_Cycle1)
        '
            If Not .DBStored Then GoTo GetEndHDate_Er
            
            .ParaNo = 4: gCnt2 = .ParaValue
            If gCnt2 <> 0 Then GoTo GetEndHDate_Er
            
            .ParaNo = 3: wHDate = .ParaValue
        End With
    Loop
    
    With CisDB
    '   <><><> 最終発注日の獲得 <><><>
        .SQL = "日付加算"
        .StoadoCount = 4
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = Input用
        .ParaNo = 3: .ParaIO = OutPut用
        .ParaNo = 4: .ParaIO = OutPut用
    '
        .ParaNo = 1: .ParaValue = wHDate
        .ParaNo = 2: .ParaValue = CisFun.Val2(iH1_Cycle1) * -1
    '
        If Not .DBStored Then GoTo GetEndHDate_Er
        
        .ParaNo = 4: gCnt2 = .ParaValue
        If gCnt2 <> 0 Then GoTo GetEndHDate_Er
        
        .ParaNo = 3: wHDate = .ParaValue
    End With
    
    iB1_SEHYmd = wHDate
    GetEndHDate = True
    
    Exit Function
GetEndHDate_Er:
    CisFun.MB_Lines = 4
    CisFun.MB_MSG(2) = "    最終発注日の取得に失敗しました            "
    CisFun.MB_Title = "最終発注日"
    CisFun.MB_Button = Error
    CisFun.MBOX
End Function

'****************************
'*      振出便設定
'****************************
Private Function FuriBin_Set() As Boolean
    FuriBin_Set = False
    
    Dim wBIdx As Byte
    Dim wBinIdx1 As Byte
    Dim wBinIdx2 As Byte
    Dim wGIdx As Byte
    Dim wBin  As Byte
    Dim w24h  As Byte
'退避ｴﾘｱｸﾘｱ
    For gIndex = 1 To BinMax + 1
        mSTTime(gIndex) = ""
        mEDTime(gIndex) = ""
        mFurBin(gIndex) = 0
        mEdaNo(gIndex) = 0
        mHCDay(gIndex) = 0
        mHCBin(gIndex) = 0
        mFurDay(gIndex) = 0
        mNYDay(gIndex) = 0
        mHYKbn(gIndex) = 0
    Next gIndex
    wBIdx = 0: w24h = 0
    
    For gIndex = 1 To BinMax
        If Trim(mBinTime(gIndex)) = "" Then Exit For
'開始･終了･振出便ｾｯﾄ
        wBIdx = wBIdx + 1
        If wBIdx = 1 Then
            '1便
            mSTTime(wBIdx) = "0000"
            mEDTime(wBIdx) = mBinTime(gIndex)
            mSTTime(wBIdx + 1) = mBinTime(gIndex)
            mEDTime(wBIdx + 1) = "2400"
            mFurBin(wBIdx) = gIndex
            mFurBin(wBIdx + 1) = gIndex
            mEdaNo(wBIdx + 1) = 1
        Else
            If mBinTime(gIndex) >= mEDTime(wBIdx - 1) Then
                '2便以降前の時間より大きい時
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
                '2便以降前の時間より小さい時
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
        '最終便の時
        If gIndex = CisFun.Val2(iH1_Cycle2) Then
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
    
'振出時間開始・終了の結果訂正
'    wYDJun = 0
    For gIndex = 1 To BinMax + 1
        If Trim(mSTTime(gIndex)) = "" Then Exit For
'--- 2004.12.21 --------------------------------------------------< Start >----
'        If Trim(mSTTime(gIndex)) = Trim(mEDTime(gIndex)) Then
'        '1便で 0000 時設定の時
'            mSTTime(gIndex) = ""
'            mEDTime(gIndex) = ""
'            mFurBin(gIndex) = 0
'        End If
'        If Trim(mSTTime(gIndex)) = "0000" And Trim(mEDTime(gIndex)) = "2400" Then
'            mEdaNo(gIndex) = 0
'        End If
'--- 2004.12.21 --------------------------------------------------< End >----
'
        If Trim(mSTTime(gIndex)) <> "" Then
            '昼夜勤退避
            mHYKbn(gIndex) = CisFun.Val2(vsGrid.TextMatrix(mFurBin(gIndex), 7))
'            If wYDJun = 0 And mHYKbn(gIndex) = 1 Then wYDJun = 1
'            If wYDJun = 1 And mHYKbn(gIndex) = 0 Then wYDJun = 2
            '納入加算日退避
            If CisFun.Val2(vsGrid.TextMatrix(mFurBin(gIndex), 3)) <> 0 Then
                mHCDay(gIndex) = CisFun.Val2(vsGrid.TextMatrix(mFurBin(gIndex), 3))
            End If
        End If

    Next gIndex
'    If wYDJun = 2 Then
'        For gIndex = 1 To BinMax + 1
'            If Trim(mSTTime(gIndex)) = "" Then Exit For
'            If mHYKbn(gIndex) = 0 Then Exit For
'            mFurDay(gIndex) = -1
'        Next gIndex
'    End If
    
    FuriBin_Set = True
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
    
    If Not HCMRead(gSL_Select, 0) Then
        Call HCMClose
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
        Call BodySet
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
Private Sub BodySet()

    With HCM
        iH1_Torcd = RTrim(.仕入先)
        iH1_Cycle1 = Left(.サイクル, 1)
        iH1_Cycle2 = Mid(.サイクル, 2, 2)
        iH1_Cycle3 = Mid(.サイクル, 4, 2)
        iH1_Ukeir = RTrim(.受入)
        
        iB1_SEHYmd = .最終発注日
        iB1_SEHBin = .最終発注便
    End With

'Grid表示
    Call GridSet
    Call HCMClose(1)
        
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
        ' 発注サイクルマスタ
        .SQL = "DELETE FROM 発注サイクルマスタ "
        .SQL = .SQL & " WHERE 仕入先 = '" & RTrim(iH1_Torcd) & "'"
        .SQL = .SQL & "   And サイクル = '" & RTrim(mCycle) & "'"
        .SQL = .SQL & "   And 受入 = '" & RTrim(iH1_Ukeir) & "'"
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
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False
    Dim AddFlg  As Boolean
    Dim wYDJun As Byte

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        If SyoriKB = "ADD" Then
            .MB_MSG(2) = "     追加処理を行います      "
            .MB_Title = "追加処理"
        Else
            .MB_MSG(2) = "     修正処理を行います      "
            .MB_Title = "修正処理"
        End If
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Call VsGridToText(Val(Input_Row.Tag))
            Call GridTextEnabled(True)
            Exit Function
        End If
    End With
    On Error GoTo DBPut_Err
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
    If SysOption = 0 Then
'システム設定【時間設定】
        If SyoriKB <> "ADD" Then
            With CisDB
                On Error GoTo DBPut_Err
        ' ( 削除処理 )
                ' 発注サイクルマスタ
                .SQL = "DELETE FROM 発注サイクルマスタ "
                .SQL = .SQL & " WHERE 仕入先 = '" & RTrim(iH1_Torcd) & "'"
                .SQL = .SQL & "   And サイクル = '" & RTrim(mCycle) & "'"
                .SQL = .SQL & "   And 受入 = '" & RTrim(iH1_Ukeir) & "'"
                .DBExec
        
                On Error GoTo 0
            End With
        End If
    '   マスタ項目移送
        For gIndex = 1 To BinMax + 1
            If RTrim(mSTTime(gIndex)) <> "" Then
                With HCM
                    .仕入先 = RTrim(iH1_Torcd)
                    .サイクル = RTrim(mCycle)
                    .受入 = RTrim(iH1_Ukeir)
                    .振出便 = Format(mFurBin(gIndex), "00")
                    .行識別 = mEdaNo(gIndex)
                    
                    .最終発注日 = Trim(iB1_SEHYmd)
                    .最終発注便 = Trim(iB1_SEHBin)
                    
                    .昼夜区分 = CisFun.Val2(vsGrid.TextMatrix(mFurBin(gIndex), 7))
                    .振出時間 = Mid(mFTime, (mFurBin(gIndex) - 1) * 4 + 1, 4)
                    .発行時間 = Mid(mHTime, (mFurBin(gIndex) - 1) * 4 + 1, 4)
                    .回収時間 = Mid(mKTime, (mFurBin(gIndex) - 1) * 4 + 1, 4)
                    .入荷時間 = Mid(mNTime, (mFurBin(gIndex) - 1) * 4 + 1, 4)
                    .振出開始 = RTrim(mSTTime(gIndex))
                    .振出終了 = RTrim(mEDTime(gIndex))
                    .振出加算日 = mFurDay(gIndex)
                    .発注加算日 = mHCDay(gIndex)
                    .発注便 = Format(mHCBin(gIndex), "00")
                    If .振出終了 > .発行時間 Then
                        If .発注加算日 = 0 Then
                            .発行加算日 = 1
                        Else
                            .発行加算日 = .発注加算日
                        End If
                    Else
                        .発行加算日 = 0 + .発注加算日
                    End If
                        
    '                .納入加算日 = mNYDay(gIndex)
                    
                End With
    '
                If SyoriKB = "ADD" Then
                    HCM.作成者 = gTanto
                Else
                    HCM.更新者 = gTanto
                End If
            ' ( 追加処理 )
                Call HCMInsert
            End If
        Next gIndex
    'シーケンス便採番(昼夜便順>夜勤は昼勤続きの便から）
        Dim wFurBin()     As Byte
        Dim wEdaNo()      As Byte
        Dim wHYKbn()      As Byte
        gSL_Select = "SELECT * FROM 発注サイクルマスタ"
        gSL_Select = gSL_Select & " WHERE 仕入先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   And サイクル = '" & RTrim(mCycle) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & " ORDER BY 振出便,行識別"
        If HCMRead(gSL_Select, 0) Then
            ReDim wFurBin(CisDB.RecordCount)
            ReDim wEdaNo(CisDB.RecordCount)
            ReDim wHYKbn(CisDB.RecordCount)
            gIndex = 0
            Do Until Not HCM_RDSTS
                gIndex = gIndex + 1
                wFurBin(gIndex) = CisFun.Val2(HCM.振出便)
                wEdaNo(gIndex) = HCM.行識別
                wHYKbn(gIndex) = HCM.昼夜区分
                Call HCMReadNext
            Loop
        End If
        Call HCMClose
        gSL_Select = "SELECT COUNT(*) 件数 FROM 発注サイクルマスタ"
        gSL_Select = gSL_Select & " WHERE 仕入先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   And サイクル = '" & RTrim(mCycle) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & "   And 昼夜区分 = 0"
        If HCMRead(gSL_Select, 1) Then
        End If
        If HCM.件数 <> 0 Then
            gCnt1 = 0
            Do Until wHYKbn(1) = 0
                wFurBin(0) = wFurBin(1)
                wEdaNo(0) = wEdaNo(1)
                wHYKbn(0) = wHYKbn(1)
                For gCnt1 = 1 To UBound(wFurBin) - 1
                    wFurBin(gCnt1) = wFurBin(gCnt1 + 1)
                    wEdaNo(gCnt1) = wEdaNo(gCnt1 + 1)
                    wHYKbn(gCnt1) = wHYKbn(gCnt1 + 1)
                Next
                wFurBin(UBound(wFurBin)) = wFurBin(0)
                wEdaNo(UBound(wFurBin)) = wEdaNo(0)
                wHYKbn(UBound(wFurBin)) = wHYKbn(0)
            Loop
        End If
        gCnt1 = 0
        For gIndex = 1 To UBound(wFurBin)
            If wEdaNo(gIndex) = 0 Then
                gCnt1 = gCnt1 + 1
            End If
            With CisDB
                On Error GoTo DBPut_Err
                .SQL = "UPDATE 発注サイクルマスタ SET"
                .SQL = .SQL & " シーケンス便 = '" & Format(gCnt1, "00") & "'"
                .SQL = .SQL & " WHERE 仕入先 = '" & RTrim(iH1_Torcd) & "'"
                .SQL = .SQL & "   And サイクル = '" & RTrim(mCycle) & "'"
                .SQL = .SQL & "   And 受入 = '" & RTrim(iH1_Ukeir) & "'"
                .SQL = .SQL & "   And 振出便 = '" & Format(wFurBin(gIndex), "00") & "'"
                .SQL = .SQL & "   And 行識別 = " & wEdaNo(gIndex) & ""
                .DBExec
        
                On Error GoTo 0
            End With
        Next
    '振出加算日の訂正(１便が夜勤から始まり２便以降昼勤が有る場合）
        gSL_Select = "SELECT * FROM 発注サイクルマスタ"
        gSL_Select = gSL_Select & " WHERE 仕入先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   And サイクル = '" & RTrim(mCycle) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & " ORDER BY 振出便,行識別"
        If HCMRead(gSL_Select, 0) Then
            ReDim wFurBin(CisDB.RecordCount)
            ReDim wEdaNo(CisDB.RecordCount)
            ReDim wHYKbn(CisDB.RecordCount)
            gIndex = 0: wYDJun = 0
            Do Until Not HCM_RDSTS
                gIndex = gIndex + 1
                wFurBin(gIndex) = CisFun.Val2(HCM.振出便)
                wEdaNo(gIndex) = HCM.行識別
                wHYKbn(gIndex) = HCM.昼夜区分
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
                    .SQL = "UPDATE 発注サイクルマスタ SET"
                    .SQL = .SQL & " 振出加算日 = 振出加算日 - 1"
                    .SQL = .SQL & " WHERE 仕入先 = '" & RTrim(iH1_Torcd) & "'"
                    .SQL = .SQL & "   And サイクル = '" & RTrim(mCycle) & "'"
                    .SQL = .SQL & "   And 受入 = '" & RTrim(iH1_Ukeir) & "'"
                    .SQL = .SQL & "   And 振出便 = '" & Format(wFurBin(gIndex), "00") & "'"
                    .SQL = .SQL & "   And 行識別 = " & wEdaNo(gIndex) & ""
                    .DBExec
            
                    On Error GoTo 0
                End With
            Next gIndex
        End If
    Else
'システム設定【発注テーブル】
        With CisDB
            On Error GoTo DBPut_Err
            .SQL = "UPDATE 発注サイクルマスタ SET"
            .SQL = .SQL & " 最終発注日 = '" & RTrim(iB1_SEHYmd) & "',"
            .SQL = .SQL & " 最終発注便 = '" & RTrim(iB1_SEHBin) & "'"
            .SQL = .SQL & " WHERE 仕入先 = '" & RTrim(iH1_Torcd) & "'"
            .SQL = .SQL & "   And サイクル = '" & RTrim(mCycle) & "'"
            .SQL = .SQL & "   And 受入 = '" & RTrim(iH1_Ukeir) & "'"
            .DBExec
    
            On Error GoTo 0
        End With
        
    End If
    
' ( 追加処理 )
    If SyoriKB = "ADD" Then
        iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
    Else
' ( 更新処理 )
        iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
    End If
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    DBPut = True
    Exit Function
DBPut_Err:
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
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 中中, あり, 6, 0, "入荷便")
        Call .FixedSet(中中, 中中, あり, 8, 1, "昼夜勤")
        Call .FixedSet(中中, 中中, あり, 8, 2, "振出時間")
        Call .FixedSet(中中, 右中, あり, 6, 3, "加算日")
        Call .FixedSet(中中, 中中, あり, 8, 4, "発行時間")
        Call .FixedSet(中中, 中中, あり, 8, 5, "回収時間")
        Call .FixedSet(中中, 中中, あり, 8, 6, "入荷時間")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 7, "x昼夜勤")
        Call .FixedSet(中中, 左中, なし, 0, 8, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6")
        .InitDisp
    End With
    
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
        If Ttl_Kensu + 1 > .Rows Then
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not HCM_RDSTS
        'Grid Row,Col
        ID = CisFun.Val2(HCM.振出便)
        With vsGrid
            .TextMatrix(ID, GridCol(0)) = ID & "便"
            If HCM.昼夜区分 = 0 Then
                .TextMatrix(ID, GridCol(1)) = "昼勤"
            Else
                .TextMatrix(ID, GridCol(1)) = "夜勤"
            End If
            .TextMatrix(ID, GridCol(2)) = Left(HCM.振出時間, 2) & ":" & Mid(HCM.振出時間, 3, 2)
            .TextMatrix(ID, GridCol(3)) = Format(HCM.発注加算日, "##")
            .TextMatrix(ID, GridCol(4)) = Left(HCM.発行時間, 2) & ":" & Mid(HCM.発行時間, 3, 2)
            If RTrim(HCM.回収時間) <> "" Then
                .TextMatrix(ID, GridCol(5)) = Left(HCM.回収時間, 2) & ":" & Mid(HCM.回収時間, 3, 2)
            End If
            If RTrim(HCM.入荷時間) <> "" Then
                .TextMatrix(ID, GridCol(6)) = Left(HCM.入荷時間, 2) & ":" & Mid(HCM.入荷時間, 3, 2)
            End If
            .TextMatrix(ID, GridCol(7)) = HCM.昼夜区分
        End With
        Call HCMReadNext
    Loop

    With vsGrid
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'==========================================================='
'   入力テキスト表示・非表示
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
    If SysOption = 0 Then
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

'テキストからグリッドに退避
    Call TextToVsGrid(Val(Input_Row.Tag))
    gInt = vsGrid.Row
    '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    '↓
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
'グリッドからテキストに表示
    Call VsGridToText(vsGrid.Row)
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    With vsGrid
        .TextMatrix(tRow, 1) = ChkB1_HYKb.Caption
        .TextMatrix(tRow, 7) = ChkB1_HYKb.Value
        
        Call EditTime(iB1_FHh, iB1_FMm, gStr)
        .TextMatrix(tRow, 2) = gStr
        If CisFun.Val2(iB1_KDd) <> 0 Then
            .TextMatrix(tRow, 3) = iB1_KDd
        Else
            .TextMatrix(tRow, 3) = ""
        End If
        Call EditTime(iB1_HHh, iB1_HMm, gStr)
        .TextMatrix(tRow, 4) = gStr
        Call EditTime(iB1_KHh, iB1_KMm, gStr)
        .TextMatrix(tRow, 5) = gStr
        Call EditTime(iB1_NHh, iB1_NMm, gStr)
        .TextMatrix(tRow, 6) = gStr
        
    End With
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With vsGrid
        Input_Row.Tag = tRow
        Input_Row.Move Input_Row.Left, vsGrid.CellTop, Input_Row.Width, vsGrid.CellHeight
        '昼夜勤
        ChkB1_HYKb.Value = CisFun.Val2((.TextMatrix(tRow, 7)))
        '振出時間
        iB1_FHh = Left(RTrim(.TextMatrix(tRow, 2)), 2)
        iB1_FMm = Mid(RTrim(.TextMatrix(tRow, 2)), 4, 2)
        '振出時間
        iB1_KDd = Format(CisFun.Val2(.TextMatrix(tRow, 3)), "##")
        '発行時間
        iB1_HHh = Left(RTrim(.TextMatrix(tRow, 4)), 2)
        iB1_HMm = Mid(RTrim(.TextMatrix(tRow, 4)), 4, 2)
        '回収時間
        iB1_KHh = Left(RTrim(.TextMatrix(tRow, 5)), 2)
        iB1_KMm = Mid(RTrim(.TextMatrix(tRow, 5)), 4, 2)
        '入荷時間
        iB1_NHh = Left(RTrim(.TextMatrix(tRow, 6)), 2)
        iB1_NMm = Mid(RTrim(.TextMatrix(tRow, 6)), 4, 2)
    End With
End Sub
Private Sub VSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
    Call TextToVsGrid(Val(Input_Row.Tag))
    Call VsGridToText(vsGrid.Row)
End Sub
Private Sub VSGrid_Click()
    Call TextToVsGrid(Val(Input_Row.Tag))
    If vsGrid.TextMatrix(vsGrid.Row, 0) = "" Then Exit Sub
    Call GridTextEnabled(True)
    Call VsGridToText(vsGrid.Row)
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub
Private Sub vsGrid_Scroll()
    'スクロール時はテキストボックスを非表示にします。
'    Call GridTextEnabled(False)
End Sub

