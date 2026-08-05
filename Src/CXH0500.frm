VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXH0500 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "入出庫・支給実績入力"
   ClientHeight    =   9945
   ClientLeft      =   135
   ClientTop       =   1500
   ClientWidth     =   14790
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
   ScaleHeight     =   9945
   ScaleWidth      =   14790
   WindowState     =   2  '最大化
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'なし
      Height          =   345
      Index           =   1
      Left            =   10260
      ScaleHeight     =   345
      ScaleWidth      =   3675
      TabIndex        =   20
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
      Height          =   795
      Left            =   6765
      ScaleHeight     =   735
      ScaleWidth      =   6075
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1620
      Width           =   6135
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   3465
         Top             =   60
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
         Left            =   4320
         Top             =   60
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
         Left            =   5175
         Top             =   60
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
         Left            =   45
         Top             =   60
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
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   585
         Left            =   900
         Top             =   60
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
         Left            =   1755
         Top             =   60
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
         Left            =   2610
         Top             =   60
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
   End
   Begin VB.PictureBox H2_Area1 
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
      Height          =   1080
      Left            =   2205
      ScaleHeight     =   1020
      ScaleWidth      =   10635
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3165
      Width           =   10695
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   915
         Left            =   6195
         Top             =   60
         Width           =   660
         _ExtentX        =   1164
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
         cPositionX      =   30
         cPositionY      =   60
         Begin CisText_V60.CisText iH2_Ukeir 
            Height          =   360
            Left            =   120
            TabIndex        =   8
            Top             =   405
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
      Begin Cis3D_v60.CIS3D Back_Torcd1 
         Height          =   915
         Left            =   2475
         Top             =   60
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "支 給 先"
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
         cPositionX      =   30
         cPositionY      =   60
         Begin CisText_V60.CisText iH2_Torcd1 
            Height          =   360
            Left            =   120
            TabIndex        =   7
            Top             =   420
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
         Begin Cis3D_v60.CIS3D H2lb_Tornm1 
            Height          =   360
            Left            =   1110
            Top             =   420
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
         Height          =   915
         Index           =   0
         Left            =   60
         Top             =   60
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "便  "
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
         cPositionX      =   60
         cPositionY      =   60
         Begin CisYMD_With_Btn.CisYMDwB iH2_Ymd 
            Height          =   360
            Left            =   30
            TabIndex        =   5
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
            Object.Height          =   360
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
            Left            =   555
            Top             =   30
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
            cPositionY      =   30
         End
         Begin CisText_V60.CisText iH2_Bin 
            Height          =   375
            Left            =   1980
            TabIndex        =   6
            Top             =   420
            Width           =   360
            _ExtentX        =   635
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
      Begin Cis3D_v60.CIS3D Back_Torcd2 
         Height          =   915
         Left            =   6855
         Top             =   60
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "支 給 元"
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
         cPositionX      =   30
         cPositionY      =   60
         Begin CisText_V60.CisText iH2_Torcd2 
            Height          =   360
            Left            =   150
            TabIndex        =   9
            Top             =   390
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
         Begin Cis3D_v60.CIS3D H2lb_Tornm2 
            Height          =   375
            Left            =   1140
            Top             =   390
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   661
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
      TabIndex        =   0
      TabStop         =   0   'False
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
      Height          =   4470
      Left            =   2205
      ScaleHeight     =   4410
      ScaleWidth      =   10635
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4695
      Width           =   10695
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   4290
         Index           =   0
         Left            =   60
         Top             =   60
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   7567
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
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   4215
            Index           =   1
            Left            =   30
            Top             =   30
            Width           =   10440
            _ExtentX        =   18415
            _ExtentY        =   7435
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
            Begin vsFlexLib.vsFlexArray vsGrid 
               Height          =   4050
               Left            =   90
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   75
               Width           =   10230
               _Version        =   196608
               _ExtentX        =   18045
               _ExtentY        =   7144
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
               FormatString    =   "№ |品　番　　　　　　　　　 |背番号  |収容数  |枚　数 |端　数 |支給数 　|有無償   |"
               Rows            =   11
               Cols            =   12
               BackColor       =   12648384
               BackColorFixed  =   12632256
               BackColorBkg    =   12632256
               FocusRect       =   0
               HighLight       =   0
               RowHeightMin    =   360
               SelectionMode   =   1
               Begin Cis3D_v60.CIS3D Input_Area 
                  Height          =   375
                  Left            =   450
                  Top             =   360
                  Width           =   9405
                  _ExtentX        =   16589
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
                     TabIndex        =   11
                     Top             =   0
                     Width           =   1065
                     _ExtentX        =   1879
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
                     Text            =   "XXXXXXXX"
                     MaxLength       =   8
                     IMEMode         =   2
                     cDataReplace    =   1
                  End
                  Begin CisText_V60.CisText iB1_Hinbn 
                     Height          =   375
                     Left            =   0
                     TabIndex        =   10
                     Top             =   0
                     Width           =   3165
                     _ExtentX        =   5583
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
                     IMEMode         =   2
                     cDataReplace    =   1
                  End
                  Begin CisText_V60.CisText iB1_Maisu 
                     Height          =   360
                     Left            =   5190
                     TabIndex        =   12
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
                     cAlignment      =   1
                  End
                  Begin CisText_V60.CisText iB1_Suryo 
                     Height          =   360
                     Left            =   7050
                     TabIndex        =   13
                     Top             =   0
                     Width           =   1200
                     _ExtentX        =   2117
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
                     cAlignment      =   1
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                     Height          =   360
                     Left            =   4140
                     Top             =   0
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   635
                     BackColor       =   16777152
                     ForeColor       =   0
                     Caption         =   "##,###"
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
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -50
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Hasu 
                     Height          =   360
                     Left            =   6120
                     Top             =   0
                     Width           =   930
                     _ExtentX        =   1640
                     _ExtentY        =   635
                     BackColor       =   16777152
                     ForeColor       =   0
                     Caption         =   "###,###"
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
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -50
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Umu 
                     Height          =   360
                     Left            =   8235
                     Top             =   0
                     Width           =   1155
                     _ExtentX        =   2037
                     _ExtentY        =   635
                     BackColor       =   14737632
                     ForeColor       =   0
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
                     cFont3DColor1   =   8421504
                     cFont3DColor2   =   -2147483643
                     cFont3D         =   2
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -50
                     Begin VB.CheckBox B1Chk_UMkbn 
                        BackColor       =   &H00E0E0E0&
                        Caption         =   "無償"
                        Height          =   255
                        Left            =   150
                        TabIndex        =   14
                        Top             =   45
                        Width           =   825
                     End
                  End
               End
               Begin Cis3D_v60.CIS3D Back_Suryo 
                  Height          =   270
                  Left            =   7290
                  Top             =   30
                  Width           =   1020
                  _ExtentX        =   1799
                  _ExtentY        =   476
                  ForeColor       =   0
                  Caption         =   "実績数"
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
                  cFont3D         =   2
                  cAlingnment     =   7
                  cBoderStyle     =   2
                  cPositionX      =   60
                  cPositionY      =   15
               End
            End
            Begin CisText_V60.CisText B_TabDummy 
               Height          =   360
               Left            =   135
               TabIndex        =   15
               Top             =   105
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
      Caption         =   "【 入出庫・支給実績入力  】"
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
         Left            =   7635
         Top             =   45
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
   Begin Cis3D_v60.CIS3D H1_Area1 
      Height          =   1020
      Left            =   2220
      Top             =   1620
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   1799
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
      Begin Cis3D_v60.CIS3D Back_NouNo 
         Height          =   840
         Left            =   2790
         Top             =   90
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1482
         ForeColor       =   16711680
         Caption         =   "納番"
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
         cPositionX      =   30
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_NouNo 
            Height          =   360
            Left            =   150
            TabIndex        =   1
            Top             =   375
            Width           =   765
            _ExtentX        =   1349
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
            cAutoEject      =   0   'False
         End
      End
      Begin Cis3D_v60.CIS3D Back_Jkbn 
         Height          =   840
         Left            =   90
         Top             =   90
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   1482
         ForeColor       =   16711680
         Caption         =   "実績区分"
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
         cPositionX      =   30
         cPositionY      =   60
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "支給"
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
            Left            =   1815
            TabIndex        =   4
            Top             =   420
            Width           =   825
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "入庫"
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
            Left            =   75
            TabIndex        =   2
            Top             =   420
            Width           =   825
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0FFC0&
            Caption         =   "出庫"
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
            Left            =   945
            TabIndex        =   3
            Top             =   420
            Value           =   -1  'True
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "CXH0500"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   入出庫・支給実績入力
'**       フォームID    :   CXH0500
'**       処理概要      :
'**
'**       作  成  日    :   2004/02/17  By CIS SIMPR-A
'**       変  更  日    :   2004/05/18  By CIS - 入力処理時、グリッド移動による不具合修正
'**       変  更  日    :   2004/05/28  By CIS - 背番号カナ入力・背番号半角チェック
'**       変  更  日    :   2004/06/23  By CIS - 1.機能追加（入出庫入力データ検索）
'**       　  　  　                             2.バグ修正（修正時、追加行の実績管理NOがｾｯﾄされない！）
'**       変  更  日    :   2004/07/06  By CIS - 1.ボディ部からヘッド部２に戻った時に処理（新規・修正・削除）ボタンが有効になってしまう！
'**       　  　  　                             2.登録処理後、ヘッド部２に戻った時に実績区分の入力制御（便・受入）が無効になってしまう！
'**       変  更  日    :   2004/12/29  By CIS - 追加更新時、納入番号=""を全て削除していた。
'**                                             　生産実績入力・出庫入力で納入番号=""が作成されるため、空白削除をなくした。
'**       変  更  日    :   2008/04/11  By CIS - 背番号８桁対応
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)
    Dim CisVsGrid       As New CisVsGrid3
    Dim ActObj          As Object
    Dim ErrObj          As Object
'   実績区分の背景色
    Dim mColor_On       As Long
    Dim mColor_Off      As Long

    Dim wTorKB          As Integer      ' 区分
    Dim wRowNo1         As Long
    Dim wRowNo2         As Long
    Dim wColNo1         As Long
    Dim wInput          As Boolean      ' 入力ﾌﾗｸﾞ
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
        .FixedHeight = 360
        .PatanMax = 1
        .SelectCol = 10
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
    SyoriKB = "ADD"
    Call SyoriName(SyoriKB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

'   背景色の設定（出力区分）
    mColor_On = H1Op_JKbn(2).BackColor
    mColor_Off = H1Op_JKbn(1).BackColor
    H1Op_JKbn(2).Value = True

    Me.Show
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
    Cancel = 1
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
   'Static W_IVENT As String
   'W_IVENT = "Down"
    Key_Acc = False

    Select Case Key_Code
           Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click    '【新規】
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    '【修正】
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
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
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
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
    Call ReturnPress
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
   Dummy.Enabled = True: Dummy.SetFocus: DoEvents
   Call CisFrm.MousePT(11)
   If ProcHB = "B1" Then
      ProcHB = "H2"
'*--------------------------------------------------<< 2004/07/06 >> Update Start
'     Call DispChange(ProcHB)
'     Call SyoriName(SyoriKB)
      Call SyoriName(SyoriKB)
      Call DispChange(ProcHB)
'*--------------------------------------------------<< 2004/07/06 >> Update End
      GoTo PB_CAN_Ed
   End If
   If ProcHB = "H2" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      Call HeadBodyClear("H2")
      Call SyoriName(SyoriKB)
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
        Dummy.Enabled = True: Dummy.SetFocus: DoEvents
        SyoriKB = "ADD"
        Call HeadBodyClear("B")
        Call HeadBodyClear("H")
        Call SyoriName(SyoriKB)
        H1Op_JKbn(2).SetFocus
        Dummy.Enabled = False
    End If
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If SyoriKB <> "MNT" Then
        Dummy.Enabled = True: Dummy.SetFocus: DoEvents
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
        iH1_NouNo.SetFocus
        Dummy.Enabled = False
    End If
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
        Dummy.Enabled = True: Dummy.SetFocus: DoEvents
        SyoriKB = "DEL"
        Call SyoriName(SyoriKB)
        iH1_NouNo.SetFocus
        Dummy.Enabled = False
    End If
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
'*--------------------------------------------------<< 2004/06/23 >> Insert Start
    If PB_Look.Tag = "iH1_NouNo" Then Call Look_JKT: Exit Sub
'*--------------------------------------------------<< 2004/06/23 >> Insert End
    If PB_Look.Tag = "iH2_Torcd1" Then Call Look_Torcd(PB_Look.Tag): Exit Sub
    If PB_Look.Tag = "iH2_Torcd2" Then Call Look_Torcd(PB_Look.Tag): Exit Sub
    If PB_Look.Tag = "iH2_Ukeir" Then Call Look_Ukeir: Exit Sub
    If PB_Look.Tag = "iB1_Hinbn" Then Call Look_Hinbn: Exit Sub
'
    If PB_Look.Tag = "iH2_Ymd" Then iH2_Ymd.ShowCalender: Exit Sub
End Sub
'+----------------------------+
'+       取引先検索            +
'+----------------------------+
Private Sub Look_Torcd(TextName As String)
    RV_Left = 0
    RV_Top = 0

    If H1Op_JKbn(3) Then
        RV_Call = "C"
        RV_TorKb = wTorKB
    '
        CKK0025.Show vbModal
        Unload CKK0025
        Set CKK0025 = Nothing
    Else
        RV_Call = "C"
        RV_TorKb = 2
    '
        CKK0020.Show vbModal
        Unload CKK0020
        Set CKK0020 = Nothing
    End If

    If RV_Rtn Then
        If TextName = "iH2_Torcd1" Then
            iH2_Torcd1 = RV_TorcdK
            H2lb_Tornm1 = RV_TorRName
        Else
            iH2_Torcd2 = RV_TorcdK
            H2lb_Tornm2 = RV_TorRName
        End If
    End If
End Sub
'+----------------------------+
'+       受入検索            +
'+----------------------------+
Private Sub Look_Ukeir()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 8            '【発注】手配先/契約先検索
    RV_Torcd = iH2_Torcd1
    RV_Kojyo = ""

    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH2_Torcd1 = RV_TorcdK
        H2lb_Tornm1 = RV_TorRName
        iH2_Ukeir = RV_Ukeir
    End If
End Sub
'+----------------------------+
'+       品番検索            +
'+----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    If H1Op_JKbn(1) Then RVI_Hinmoku = 9
    If H1Op_JKbn(2) Then RVI_Hinmoku = 9
    If H1Op_JKbn(3) Then RVI_Hinmoku = 1

    CKK0030.Show vbModal
    Unload CKK0030
    Set CKK0030 = Nothing

    If RV_Rtn Then
       iB1_Hinbn = RTrim(RV_Hinbn)
    End If
End Sub
'*--------------------------------------------------<< 2004/06/23 >> Insert Start
'+-------------------------------------+
'+       入出庫入力データ検索
'+-------------------------------------+
Private Sub Look_JKT()
    RV_Left = 0
    RV_Top = 0
    RV_Call = ""

    CXH0501.Show vbModal
    Unload CXH0501
    Set CXH0501 = Nothing

    If RV_Rtn Then
       iH1_NouNo = RTrim(RV_Code)
    End If
End Sub
'*--------------------------------------------------<< 2004/06/23 >> Insert End
'+-----------------------------+
'+      H1Op_JKbn
'+-----------------------------+
Private Sub H1Op_JKbn_Click(Index As Integer)
'   背景色
    For gInt = 1 To 3
        If gInt = Index Then
            H1Op_JKbn(gInt).BackColor = mColor_On
        Else
            H1Op_JKbn(gInt).BackColor = mColor_Off
        End If
    Next gInt
'   項目名＆入力制御
    Select Case Index
        Case 1
        '   <><><> 入庫 <><><>
            wTorKB = 0          '【全て】
            Back_Ymd.Caption = "入 庫 日"
            Back_Torcd1.Caption = "入 庫 先"
            Back_Torcd2.Caption = "入 庫 元"
            Back_Suryo.Caption = "入庫数"
            iH2_Bin.Enabled = False
            iH2_Ukeir.Enabled = False
            B1Chk_UMkbn.Enabled = False
        Case 2
        '   <><><> 出庫 <><><>
            wTorKB = 0          '【全て】
            Back_Ymd.Caption = "出 庫 日"
            Back_Torcd1.Caption = "出 庫 先"
            Back_Torcd2.Caption = "出 庫 元"
            Back_Suryo.Caption = "出庫数"
            iH2_Bin.Enabled = False
            iH2_Ukeir.Enabled = False
            B1Chk_UMkbn.Enabled = False
        Case 3
        '   <><><> 支給 <><><>
            wTorKB = 8          '【発注】手配先/契約先検索
            Back_Ymd.Caption = "納 入 日"
            Back_Torcd1.Caption = "支 給 先"
            Back_Torcd2.Caption = "支 給 元"
            Back_Suryo.Caption = "支給数"
            iH2_Bin.Enabled = True
            iH2_Ukeir.Enabled = True
            B1Chk_UMkbn.Enabled = True
    End Select
End Sub
'*--------------------------------------------------<< 2004/06/23 >> Insert Start
'+-----------------------------+
'+      iH1_NouNo
'+-----------------------------+
Private Sub iH1_NouNo_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_NouNo_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'*--------------------------------------------------<< 2004/06/23 >> Insert End
'+-----------------------------+
'+      iH2_Ymd
'+-----------------------------+
Private Sub iH2_Ymd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH2_Ymd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+      iH2_Torcd1
'+-----------------------------+
Private Sub iH2_Torcd1_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH2_Torcd1_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H2lb_Tornm1 = ""
    If RTrim(iH2_Torcd1) <> "" Then
        If Get_Torcd(iH2_Torcd1, wTorKB) Then
            H2lb_Tornm1 = TRM_RName
        End If
    End If
End Sub
'+-----------------------------+
'+      iH2_Ukeir
'+-----------------------------+
Private Sub iH2_Ukeir_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH2_Ukeir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+      iH2_Torcd2
'+-----------------------------+
Private Sub iH2_Torcd2_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH2_Torcd2_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H2lb_Tornm2 = ""
    If RTrim(iH2_Torcd2) <> "" Then
        If Get_Torcd(iH2_Torcd2, wTorKB) Then
            H2lb_Tornm2 = TRM_RName
        End If
    End If
End Sub
'+-----------------------------+
'+      iB1_Hinbn
'+-----------------------------+
Private Sub iB1_Hinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Set ActObj = iB1_Hinbn
    B_TabDummy.Tag = "S"
End Sub
Private Sub iB1_Hinbn_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    Call Hinbn_Set
End Sub
'+-----------------------------+
'+      iB1_Seban
'+-----------------------------+
Private Sub iB1_Seban_GotFocus()
    Set ActObj = iB1_Seban
End Sub
Private Sub iB1_Seban_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
'+-----------------------------+
'+      iB1_Maisu
'+-----------------------------+
Private Sub iB1_Maisu_GotFocus()
    Set ActObj = iB1_Maisu
End Sub
Private Sub iB1_Maisu_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Maisu_KeyUp(KeyCode As Integer, Shift As Integer)
'   最大桁数のﾁｪｯｸ
    Select Case KeyCode
        Case 48 To 57, 96 To 105
            If CisFun.Val2(B1lb_Syuyo) <> 0 Then
                gLong = CisFun.Val2(B1lb_Syuyo) * CisFun.Val2(iB1_Maisu & "9")
                If iB1_Suryo.MaxLength < CisFun.Len2(Format(gLong, "####")) Then
                    SendKeys "{TAB}"
                End If
            End If
    End Select
End Sub
Private Sub iB1_Maisu_LostFocus()
    If CisFun.Val2(iB1_Maisu) <> 0 Then
        If CisFun.Val2(B1lb_Syuyo) <> 0 Then
            gLong = (CisFun.Val2(B1lb_Syuyo) * CisFun.Val2(iB1_Maisu))
            iB1_Suryo = Format(gLong, "####")
        End If
    End If
    B1lb_Hasu = ""
End Sub
'+-----------------------------+
'+      iB1_Suryo
'+-----------------------------+
Private Sub iB1_Suryo_GotFocus()
    Set ActObj = iB1_Suryo
    B_TabDummy.Tag = "E"
End Sub
Private Sub iB1_Suryo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Suryo_LostFocus()
    Call Suryo_Set
End Sub
'+-----------------------------+
'+      B1Chk_UMkbn
'+-----------------------------+
Private Sub B1Chk_UMkbn_GotFocus()
    Set ActObj = B1Chk_UMkbn
    B_TabDummy.Tag = "E"
End Sub
Private Sub B1Chk_UMkbn_Click()
    Select Case B1Chk_UMkbn.Value
        Case 0: B1Chk_UMkbn.Caption = "無償"
        Case 1: B1Chk_UMkbn.Caption = "有償"
    End Select
End Sub
Private Sub B1Chk_UMkbn_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
'+-----------------------------+
'+      B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
    If B_TabDummy.Tag = "S" Then
        Call TextToVsGrid(Val(Input_Area.Tag))
        Call Grid_RowChange(vbKeyUp)
    End If
    If B_TabDummy.Tag = "E" Then
        Call TextToVsGrid(Val(Input_Area.Tag))
        Call Grid_RowChange(vbKeyDown)
    End If
    SendKeys "{Tab}"
End Sub

'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
    Dim w_Color     As Long
    w_Color = &H80000008 '初期色ｾｯﾄ(黒)
    Select Case SYKB
        Case "ADD"
            SyoriNM = "追加"
            K_Sykbnm.ForeColor = &HFFFF&   '(黄色)
        Case "MNT"
            SyoriNM = "修正"
            K_Sykbnm.ForeColor = &HFF0000  '(青色)
        Case "DEL"
            SyoriNM = "削除"
            K_Sykbnm.ForeColor = &HFF&     '(赤色)
            w_Color = &HFF&
        Case "LOK"
            SyoriNM = "照会"
            K_Sykbnm.ForeColor = &HFF00FF   '(ピンク)
        Case Else
            SyoriNM = ""
    End Select
    For Each gObj In Me.Controls
        If gObj.Name Like "i*" Then gObj.ForeColor = w_Color
        If gObj.Name Like "*Comb*" Then gObj.ForeColor = w_Color
        If gObj.Name Like "*vsGrid*" Then gObj.ForeColor = w_Color
'       * PF-Key 表示設定 *
        If SYKB <> "LOC" Then
            If gObj.Name = "PB_ADD" Or _
               gObj.Name = "PB_MNT" Or _
               gObj.Name = "PB_DEL" Then
               gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SYKB)
            End If
        End If
    Next gObj
'   // 入力制御 //
    Select Case SYKB
        Case "ADD"
            Back_Jkbn.Enabled = True
            Back_NouNo.Enabled = False
            iH1_NouNo.cLostColor = &HFFFFC0
        Case Else
            Back_Jkbn.Enabled = False
            Back_NouNo.Enabled = True
            iH1_NouNo.cLostColor = &HC0E0FF
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
    Dim H2Mode      As Boolean
    Dim H2Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long

    H1Mode = False: H1Color = gPLostSel
    H2Mode = False: H2Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1":   H1Mode = True:  H1Color = gPGotSel
       Case "H2":   H2Mode = True:  H2Color = gPGotSel
       Case "B1":   B1Mode = True:  B1Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode:  H1_Area1.BackColor = H1Color
    H2_Area1.Enabled = H2Mode:  H2_Area1.BackColor = H2Color
    B1_Area1.Enabled = B1Mode:  B1_Area1.BackColor = B1Color

    PB_Look.Visible = False
    PB_ADD.Visible = False
    PB_MNT.Visible = False
    PB_DEL.Visible = False
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode

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
'*--------------------------------------------------<< 2004/07/06 >> Update Start
'           If .Name Like "i" & wbc_Enm & "_*" Then .Text = "": .Enabled = True
            If .Name Like "i" & wbc_Enm & "_*" Then .Text = ""
'*--------------------------------------------------<< 2004/07/06 >> Update End
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False: .BackColor = mColor_Off
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Cmb_*" Then .Clear
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
'       << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call HeadBodyClear("B")
            Call HeadBodyClear("H2")
            ProcHB = "H1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
        Else
            ProcHB = "H2"
            Call DispChange(ProcHB)
        End If
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "H2" Then
        If Not Head2Chk Then GoTo ReturnPress_Ed
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
        Call HeadBodyClear("H2")
        Select Case SyoriKB
            Case "ADD"
                ProcHB = "H2"
                Call DispChange(ProcHB)
            Case Else
                ProcHB = "H1"
                Call DispChange(ProcHB)
                Call SyoriName(SyoriKB)
        End Select
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

    If SyoriKB = "ADD" Then GoTo Head1Chk_Ed

    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
'   納番
    If RTrim(iH1_NouNo) = "" Then
        CisFun.MB_MSG(2) = "    納番を入力して下さい            "
        CisFun.MBOX
        iH1_NouNo.SetFocus
        Exit Function
    End If

' マスタ入力 SELECT文 ｾｯﾄ
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  jkt.*"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  実績管理テーブル jkt"
    gSL_Select = gSL_Select & " where jkt.納入番号 = '" & RTrim(iH1_NouNo) & "'"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  jkt.納入番号"
    gSL_Select = gSL_Select & ", jkt.納品書行"
    If Not DBInput("INV") Then
        CisFun.MB_MSG(2) = "    実績データ　未登録！            "
        CisFun.MBOX
        Exit Function
    End If

Head1Chk_Ed:
    Head1Chk = True
End Function
'****************************
'*      HEAD2ﾁｪｯｸ処理
'****************************
Private Function Head2Chk() As Boolean
    Head2Chk = False

    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
'   入庫日 or 出庫日 or 納入日
    If RTrim(iH2_Ymd) = "" Then
        CisFun.MB_MSG(2) = "    " & Back_Ymd.Caption & "を入力して下さい            "
        CisFun.MBOX
        iH2_Ymd.SetFocus
        Exit Function
    Else
        If Not iH2_Ymd.cChkResult Then
            CisFun.MB_MSG(2) = "    " & Back_Ymd.Caption & "を正しく入力して下さい            "
            CisFun.MBOX
            iH2_Ymd.SetFocus
            Exit Function
        End If
    End If
'   納入便
    If RTrim(iH2_Bin) = "" Then
        If H1Op_JKbn(3) Then
            CisFun.MB_MSG(2) = "    納入便を入力して下さい            "
            CisFun.MBOX
            iH2_Bin.SetFocus
            Exit Function
        End If
    End If
'   入庫先 or 出庫先 or 支給先
    If Trim(iH2_Torcd1) = "" Then
        If H1Op_JKbn(3) Then
            CisFun.MB_MSG(2) = "    " & Back_Torcd1.Caption & "を入力して下さい            "
            CisFun.MBOX
            iH2_Torcd1.SetFocus
            Exit Function
        End If
    Else
        If Not Get_Torcd(iH2_Torcd1, wTorKB) Then
            CisFun.MB_MSG(2) = "    " & Back_Torcd1.Caption & "　未登録！（取引先マスタ）            "
            CisFun.MBOX
            iH2_Torcd1.SetFocus
            Exit Function
        End If
    End If
'   受入
    If RTrim(iH2_Ukeir) = "" Then
        If H1Op_JKbn(3) Then
            CisFun.MB_MSG(2) = "    受入を入力して下さい            "
            CisFun.MBOX
            iH2_Ukeir.SetFocus
            Exit Function
        End If
    Else
        TUM.取引先CD = RTrim(iH2_Torcd1)
        TUM.受入 = RTrim(iH2_Ukeir)
        TUM.取引先区分 = 1
        If Not TUMRead("", 1) Then
            CisFun.MB_MSG(2) = "    受入　未登録！（取引先受入マスタ）            "
            CisFun.MBOX
            iH2_Ukeir.SetFocus
            Exit Function
        End If
    End If
'   入庫元 or 出庫元 or 支給元
    If Trim(iH2_Torcd2) = "" Then
        If H1Op_JKbn(3) Then
            CisFun.MB_MSG(2) = "    " & Back_Torcd2.Caption & "を入力して下さい            "
            CisFun.MBOX
            iH2_Torcd2.SetFocus
            Exit Function
        End If
    Else
        If Not Get_Torcd(iH2_Torcd2, wTorKB) Then
            CisFun.MB_MSG(2) = "    " & Back_Torcd2.Caption & "　未登録！（取引先マスタ）            "
            CisFun.MBOX
            iH2_Torcd2.SetFocus
            Exit Function
        End If
    End If

    With vsGrid
        .Redraw = False
        For gInt = 1 To 10
            'No
            .TextMatrix(gInt, CisVsGrid.FixedGet(0)) = gInt
        Next
        .Redraw = True
        .Row = 1
    End With

    Head2Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False

    Call TextToVsGrid(Val(Input_Area.Tag))
    Call GridTextEnabled(False)

'ﾒｯｾｰｼﾞﾎﾞｯｸｽ設定
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
'
    With vsGrid
'       << 行を詰める >>
        wInput = False
        For wRowNo1 = 1 To .Rows - 1
            If RTrim(.TextMatrix(wRowNo1, 1)) = "" Then
                For wRowNo2 = wRowNo1 + 1 To .Rows - 1
                    If RTrim(.TextMatrix(wRowNo2, 1)) <> "" Then
                        For wColNo1 = 1 To 9
                            .TextMatrix(wRowNo1, wColNo1) = .TextMatrix(wRowNo2, wColNo1)
                            .TextMatrix(wRowNo2, wColNo1) = ""
                        Next wColNo1
'------------------------------------------------------------<< 2004/05/18 >> Insert Start
                        wInput = True
                        '> 行内容を移動時、入力域の再セット
                        If Val(Input_Area.Tag) = wRowNo1 Or _
                           Val(Input_Area.Tag) = wRowNo2 Then
                            Call VsGridToText(Val(Input_Area.Tag))
                        End If
'------------------------------------------------------------<< 2004/05/18 >> Insert End
                        Exit For
                    End If
                Next wRowNo2
            Else
                wInput = True
            End If
        Next wRowNo1
        If Not wInput Then
            CisFun.MB_MSG(2) = "    明細を入力して下さい            "
            CisFun.MBOX
            Set ErrObj = iB1_Hinbn
            wRowNo1 = 1
            GoTo Body1Chk_Err
        End If
'
        For wRowNo1 = 1 To .Rows - 1
            If RTrim(.TextMatrix(wRowNo1, 1)) = "" Then Exit For

        '   << 重複チェック >>
            For wRowNo2 = wRowNo1 + 1 To .Rows - 1
                If RTrim(.TextMatrix(wRowNo1, 1)) = RTrim(.TextMatrix(wRowNo2, 1)) Then
                    CisFun.MB_MSG(1) = "    品番が重複しています。            "
                    CisFun.MB_MSG(3) = "    " & wRowNo1 & "行目と " & wRowNo2 & "行目            "
                    CisFun.MBOX
                    Set ErrObj = iB1_Hinbn
                    wRowNo1 = wRowNo2
                    GoTo Body1Chk_Err
                End If
            Next wRowNo2
        '   << 品番マスタ存在チェック >>
            HNM.品番 = RTrim(.TextMatrix(wRowNo1, 1))
            If Not HNMRead("", 1) Then
                CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　品番　未登録！（品番マスタ）            "
                CisFun.MBOX
                Set ErrObj = iB1_Hinbn
                GoTo Body1Chk_Err
            End If
        ' 背番号半角チェック
            If RTrim(.TextMatrix(wRowNo1, 2)) <> "" And Not HalfSizeChk(.TextMatrix(wRowNo1, 2)) Then
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
        '   << 実績数チェック >>
            If CisFun.Val2(.TextMatrix(wRowNo1, 6)) = 0 Then
                CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　" & Back_Suryo.Caption & "を入力して下さい。            "
                CisFun.MBOX
                Set ErrObj = iB1_Suryo
                GoTo Body1Chk_Err
            End If
'
        Next wRowNo1
    End With
'
    Body1Chk = True
    Exit Function
Body1Chk_Err:
    vsGrid.Row = wRowNo1
    Call VsGridToText(vsGrid.Row)
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
    On Error GoTo DBInput_Err:

    If GridSet Then
        With JKT
            iH1_NouNo = RTrim(.納入番号)
            Select Case .実績区分
                Case 1: H1Op_JKbn(1).Value = True
                Case 2: H1Op_JKbn(2).Value = True
                Case 3: H1Op_JKbn(3).Value = True
            End Select
            iH2_Ymd = .実績日
            iH2_Bin = .実績便
            iH2_Torcd1 = RTrim(.実績先)
            iH2_Torcd2 = RTrim(.実績元)
            iH2_Ukeir = RTrim(.受入)
            If Get_Torcd(.実績先, 0) Then H2lb_Tornm1 = RTrim(TRM.略称)
            If Get_Torcd(.実績元, 0) Then H2lb_Tornm2 = RTrim(TRM.略称)
        End With
        DBInput = True
    End If

DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
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

' ( 確認メッセージ )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "    削除処理を行います。            "
        .MB_MSG(4) = "     よろしいですか？            "
        .MB_Title = "削除確認"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With

    On Error GoTo DBDelete_Err
    Call CisDB.DBTran(TransBegin)   ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

' ( 削除処理 )
    Call DBDelete_JKT
'   件数表示
    iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")

    Call CisDB.DBTran(TransCommit)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    On Error GoTo 0

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

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True

' ( 確認メッセージ )
    With CisFun
        .MB_Lines = 5
        If SyoriKB = "ADD" Then
            .MB_MSG(2) = "    追加処理を行います。            "
            .MB_Title = "追加確認"
        Else
            .MB_MSG(2) = "    修正処理を行います。            "
            .MB_Title = "修正確認"
        End If
        .MB_MSG(4) = "     よろしいですか？            "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Call VsGridToText(Val(Input_Area.Tag))
            Call GridTextEnabled(True)
            Exit Function
        End If
    End With

    On Error GoTo DBPut_Err
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
' ( 削除処理 )
'---<< 2004.12.30 追加更新時の削除はなし  >>---------------------------Start------------
    If SyoriKB = "MNT" Then
        Call DBDelete_JKT
    End If
'---<< 2004.12.30 追加更新時の削除はなし  >>---------------------------End  ------------
' ( 追加処理 )
    If Not DBSet_JKT Then GoTo DBPut_Err
    For gIndex = 1 To vsGrid.Rows - 1
        If RTrim(vsGrid.TextMatrix(gIndex, 1)) <> "" Then
            Call DBInsert_JKT(gIndex)
        End If
    Next gIndex
'   件数表示
    Select Case SyoriKB
        Case "ADD": iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
        Case "MNT": iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
    End Select

    Call CisDB.DBTran(TransCommit)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    On Error GoTo 0

    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+------------------------------------+
'+      削除処理 - 実績管理テーブル
'+------------------------------------+
Private Function DBDelete_JKT(Optional DBNo As Integer) As Boolean
    DBDelete_JKT = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "delete 実績管理テーブル"
    gSL_Select = gSL_Select & " where 納入番号 = '" & RTrim(iH1_NouNo) & "'"
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With

    DBDelete_JKT = True
End Function
'+------------------------------------+
'+      セット - 実績管理テーブル
'+------------------------------------+
Private Function DBSet_JKT()
    DBSet_JKT = False

    Call ItemsClearJKT
    With JKT
        If H1Op_JKbn(1) Then .実績区分 = 1
        If H1Op_JKbn(2) Then .実績区分 = 2
        If H1Op_JKbn(3) Then .実績区分 = 3
        .実績先 = RTrim(iH2_Torcd1)
        .受入 = RTrim(iH2_Ukeir)
        .実績元 = RTrim(iH2_Torcd2)
        .実績日 = RTrim(iH2_Ymd)
        .実績便 = RTrim(iH2_Bin)
        .作成者 = gTanto
    End With

    If SyoriKB = "ADD" Then
'       ## 追加の場合 ##
        With CisDB
        '   <><><> 納入番号の獲得 <><><>
            .SQL = "実績納入番号獲得"
            .StoadoCount = 2
            .ParaNo = 0: .ParaIO = Return用
            .ParaNo = 1: .ParaIO = OutPut用
            .ParaNo = 2: .ParaIO = OutPut用
        '
            If Not .DBStored Then
                CisFun.MB_Lines = 4
                CisFun.MB_MSG(2) = "    実績納入番号の獲得に失敗しました            "
                CisFun.MB_Title = "採番処理"
                CisFun.MB_Button = Error
                CisFun.MBOX
                Exit Function
            Else
                .ParaNo = 1: JKT.納入番号 = .ParaValue & "0"
            End If
        End With
    Else
'       ## 修正の場合 ##
        JKT.納入番号 = RTrim(iH1_NouNo)
    End If

    DBSet_JKT = True
End Function
'+------------------------------------+
'+      追加処理 - 実績管理テーブル
'+------------------------------------+
Private Function DBInsert_JKT(ID As Long, Optional DBNo As Integer) As Boolean
    DBInsert_JKT = False

'*--------------------------------------------------<< 2004/06/23 >> Update Start
'   If SyoriKB = "ADD" Then
    If CisFun.Val2(vsGrid.TextMatrix(ID, 9)) = 0 Then
'*--------------------------------------------------<< 2004/06/23 >> Update End
'       ## 追加の場合 ##
        With CisDB
        '   <><><> 実績管理番号の獲得 <><><>
            .SQL = "番号獲得"
            .StoadoCount = 3
            .ParaNo = 0: .ParaIO = Return用
            .ParaNo = 1: .ParaIO = Input用
            .ParaNo = 2: .ParaIO = OutPut用
            .ParaNo = 3: .ParaIO = Return用
        '
            .ParaNo = 1: .ParaValue = 6
        '
            If Not .DBStored Then
                CisFun.MB_Lines = 4
                CisFun.MB_MSG(2) = "    実績管理番号の獲得に失敗しました            "
                CisFun.MB_Title = "採番処理"
                CisFun.MB_Button = Error
                CisFun.MBOX
                Exit Function
            Else
                .ParaNo = 2: JKT.実績管理NO = .ParaValue
            End If
        End With
    Else
'       ## 修正の場合 ##
        JKT.実績管理NO = CisFun.Val2(vsGrid.TextMatrix(ID, 9))
    End If

    With JKT
        .納品書行 = ID
        .品番 = RTrim(vsGrid.TextMatrix(ID, 1))
        .背番号 = RTrim(vsGrid.TextMatrix(ID, 2))
        .収容数 = CisFun.Val2(vsGrid.TextMatrix(ID, 3))
        .実績枚数 = CisFun.Val2(vsGrid.TextMatrix(ID, 4))
        .実績端数 = CisFun.Val2(vsGrid.TextMatrix(ID, 5))
        .実績数 = CisFun.Val2(vsGrid.TextMatrix(ID, 6))
        .有無償区分 = CisFun.Val2(vsGrid.TextMatrix(ID, 8))
    End With
    Call JKTInsert

    DBInsert_JKT = True
End Function
'
'*------------------------------------------*'
'*    品番による背番号・収容数の表示          *
'*------------------------------------------*'
Private Sub Hinbn_Set()
    If RTrim(iB1_Hinbn) = "" Then Exit Sub

    gSL_Select = ""
    gSL_Select = gSL_Select & "select *"
    gSL_Select = gSL_Select & "  from 品番取引先マスタ"
    gSL_Select = gSL_Select & " where 品番 = '" & RTrim(iB1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   and 取引先 = '" & RTrim(iH2_Torcd1) & "'"
    gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(iH2_Ukeir) & "'"
    If HTMRead(gSL_Select, 1) Then
        iB1_Seban = RTrim(HTM.背番号)
'       収容数表示
        Call Syuyo_Set
    Else
        gSL_Select = ""
        gSL_Select = gSL_Select & "select *"
        gSL_Select = gSL_Select & "  from 品番マスタ"
        gSL_Select = gSL_Select & " where 品番 = '" & RTrim(iB1_Hinbn) & "'"
        If HNMRead(gSL_Select, 1) Then
            B1lb_Syuyo = Format(HNM.基本収容数, "#,###")
            Call Suryo_Set
        End If
    End If
End Sub
'*------------------------------------------*'
'*    収容数の表示                           *
'*------------------------------------------*'
Private Sub Syuyo_Set()
    If RTrim(HTM.変更日) <> "" Then
        If RTrim(iH2_Bin) <> "" Then
            If RTrim(HTM.変更日) & RTrim(HTM.変更便) > RTrim(iH2_Ymd) & RTrim(iH2_Bin) Then
                B1lb_Syuyo = Format(HTM.収容数1, "#,###")
            Else
                B1lb_Syuyo = Format(HTM.収容数2, "#,###")
            End If
        Else
            If RTrim(HTM.変更日) > RTrim(iH2_Ymd) Then
                B1lb_Syuyo = Format(HTM.収容数1, "#,###")
            Else
                B1lb_Syuyo = Format(HTM.収容数2, "#,###")
            End If
        End If
    Else
        B1lb_Syuyo = Format(HTM.収容数1, "#,###")
    End If
    Call Suryo_Set
End Sub
'*------------------------------------------*'
'*    枚数・端数・実績数の再表示              *
'*------------------------------------------*'
Private Sub Suryo_Set()
    B1lb_Hasu = ""
    If CisFun.Val2(B1lb_Syuyo) <> 0 Then
        If RTrim(iB1_Suryo) = "" Then
            If CisFun.Val2(iB1_Maisu) <> 0 Then
                gLong = (CisFun.Val2(B1lb_Syuyo) * CisFun.Val2(iB1_Maisu))
                iB1_Suryo = Format(gLong, "####")
            End If
        Else
            gCnt1 = CisFun.Val2(iB1_Suryo) \ CisFun.Val2(B1lb_Syuyo)
            gCnt2 = CisFun.Val2(iB1_Suryo) Mod CisFun.Val2(B1lb_Syuyo)
            If gCnt2 <> 0 Then
                gCnt1 = gCnt1 + 1
            End If
            iB1_Maisu = Format(gCnt1, "####")
            B1lb_Hasu = Format(gCnt2, "#,###")
        End If
    Else
        iB1_Maisu = ""
        B1lb_Hasu = ""
    End If
End Sub



'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理　関　連
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 左中, あり, 25, 1, "品　番")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
        'Call .FixedSet(中中, 左中, あり, 6, 2, "背番号")
        Call .FixedSet(中中, 左中, あり, 8, 2, "背番号")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
        Call .FixedSet(中中, 右中, あり, 8, 3, "収容数")
        Call .FixedSet(中中, 右中, あり, 7, 4, "枚　数")
        Call .FixedSet(中中, 右中, あり, 7, 5, "端　数")
        Call .FixedSet(中中, 右中, あり, 9, 6, "支給数")
        Call .FixedSet(中中, 中中, なし, 9, 7, "有無償")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 8, "有無償区分")
        Call .FixedSet(中中, 左中, なし, 0, 9, "実績管理NO")
        Call .FixedSet(中中, 左中, なし, 0, 10, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7")
        .InitDisp
    End With
    
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Function GridSet() As Boolean
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

    GridSet = False

    If JKTRead(gSL_Select, , 1) Then

'       抽出処理 初期ｾｯﾄ
        Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
'       ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
        With vsGrid
           .Redraw = False
            If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
                .Rows = CisVsGrid.DispMax
            Else
                .Rows = Ttl_Kensu + 1
            End If
        End With
        ID = 0
        Do Until Not JKT_RDSTS
            ID = ID + 1

            Call JktInfoToGrid(ID)

            Call JKTReadNext(1)
        Loop
        With vsGrid
             .Row = 1: .Col = 1: .ColSel = .Cols - 1
             .Redraw = True
        End With

        GridSet = True
    End If
    Call JKTClose
End Function
'+----------------------------------------------+
'+      実績管理テーブルよりグリッドに転送
'+----------------------------------------------+
Private Sub JktInfoToGrid(SetRow As Long)
    With vsGrid
        'No
        .TextMatrix(SetRow, 0) = SetRow
        '品番
        .TextMatrix(SetRow, 1) = RTrim(JKT.品番)
        '背番号
        .TextMatrix(SetRow, 2) = RTrim(JKT.背番号)
        '収容数
        .TextMatrix(SetRow, 3) = Format(JKT.収容数, "#,###")
        '枚数
        .TextMatrix(SetRow, 4) = Format(JKT.実績枚数, "#,###")
        '端数
        .TextMatrix(SetRow, 5) = Format(JKT.実績端数, "#,###")
        '実績数
        .TextMatrix(SetRow, 6) = Format(JKT.実績数, "#,###")
        '有無償
        If JKT.実績区分 = 3 Then
            Select Case JKT.有無償区分
                Case 0: .TextMatrix(SetRow, 7) = "無償"
                Case 1: .TextMatrix(SetRow, 7) = "有償"
            End Select
        End If
        .TextMatrix(SetRow, 8) = JKT.有無償区分
        .TextMatrix(SetRow, 9) = JKT.実績管理NO
    End With
End Sub
'==========================================================='
'   入力テキスト表示・非表示
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
'------------------------------------------------------------<< 2004/05/18 >> Update Start
'   Input_Area.Visible = EnabledFlg
'   B_TabDummy.Visible = EnabledFlg
'   *** ↑: 上記の１行は、入力処理時にグリッド移動が発生しないようにする為！ ***
'   *** カーソル位置が移動しない様にする為！ ***
    If EnabledFlg Then
        Input_Area.Visible = EnabledFlg
        B_TabDummy.Visible = EnabledFlg
    Else
        B_TabDummy.Visible = EnabledFlg
        Input_Area.Visible = EnabledFlg
    End If
'------------------------------------------------------------<< 2004/05/18 >> Update End
End Sub
'==========================================================='
'   グリッド移動
'==========================================================='
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And _
       KeyCode <> vbKeyDown And _
       KeyCode <> vbKeyPageUp And _
       KeyCode <> vbKeyPageDown Then Exit Sub

'   テキストからグリッドに退避
    Call TextToVsGrid(Val(Input_Area.Tag))
    gInt = vsGrid.Row
   '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
   '↓
    If KeyCode = vbKeyDown Then
        If gInt + 1 > vsGrid.Rows - 1 Then
            If RTrim(vsGrid.TextMatrix(gInt, 1)) <> "" Then
                vsGrid.Rows = gInt + 1 + 1
                vsGrid.TextMatrix(gInt + 1, 0) = gInt + 1
            Else
                GoTo Grid_RowChange_End
            End If
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
    If KeyCode = vbKeyPageDown Then
        If gInt + CisVsGrid.DispMax > vsGrid.Rows - 1 Then
            gInt = vsGrid.Rows - 1
        Else
            gInt = gInt + CisVsGrid.DispMax
        End If
    End If
    vsGrid.Row = gInt
'   グリッドからテキストに表示
    Call VsGridToText(vsGrid.Row)
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(RowNo As Integer)
    Call Suryo_Set
    With vsGrid
        '品番
        .TextMatrix(RowNo, 1) = RTrim(iB1_Hinbn)
        '背番号
        .TextMatrix(RowNo, 2) = RTrim(iB1_Seban)
        '収容数
        .TextMatrix(RowNo, 3) = Format(CisFun.Val2(B1lb_Syuyo), "#,###")
        '枚数
        If CisFun.Val2(iB1_Maisu) <> 0 Then
            .TextMatrix(RowNo, 4) = Format(CisFun.Val2(iB1_Maisu), "#,###")
        Else
            .TextMatrix(RowNo, 4) = ""
        End If
        '端数
        .TextMatrix(RowNo, 5) = Format(CisFun.Val2(B1lb_Hasu), "#,###")
        '実績数
        If CisFun.Val2(iB1_Suryo) <> 0 Then
            .TextMatrix(RowNo, 6) = Format(CisFun.Val2(iB1_Suryo), "#,###")
        Else
            .TextMatrix(RowNo, 6) = ""
        End If
        '有無償
        .TextMatrix(RowNo, 7) = ""
        If RTrim(iB1_Hinbn) <> "" Then
            If H1Op_JKbn(3) Then .TextMatrix(RowNo, 7) = B1Chk_UMkbn.Caption
        End If
        .TextMatrix(RowNo, 8) = B1Chk_UMkbn.Value
    End With
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(RowNo As Integer)
    With vsGrid
        Input_Area.Tag = RowNo
        Input_Area.Move Input_Area.Left, vsGrid.CellTop, Input_Area.Width, vsGrid.CellHeight
        '品番
        iB1_Hinbn = RTrim(.TextMatrix(RowNo, 1))
        '背番号
        iB1_Seban = RTrim(.TextMatrix(RowNo, 2))
        '収容数
        B1lb_Syuyo = Format(CisFun.Val2(.TextMatrix(RowNo, 3)), "#,###")
        '枚数
        iB1_Maisu = CisFun.Val2(.TextMatrix(RowNo, 4))
        '端数
        B1lb_Hasu = Format(CisFun.Val2(.TextMatrix(RowNo, 5)), "#,###")
        '支給数
        iB1_Suryo = CisFun.Val2(.TextMatrix(RowNo, 6))
        '有無償
        B1Chk_UMkbn = CisFun.Val2(.TextMatrix(RowNo, 8))
    End With
End Sub
Private Sub VSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And _
       KeyCode <> vbKeyDown And _
       KeyCode <> vbKeyPageUp And _
       KeyCode <> vbKeyPageDown Then Exit Sub
    Call TextToVsGrid(Val(Input_Area.Tag))
    Call VsGridToText(vsGrid.Row)
End Sub
Private Sub VSGrid_Click()
    Call TextToVsGrid(Val(Input_Area.Tag))
    If vsGrid.TextMatrix(vsGrid.Row, 0) = "" Then Exit Sub
    Call GridTextEnabled(True)
    Call VsGridToText(vsGrid.Row)
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub
Private Sub VSGrid_DblClick()
    If RTrim(vsGrid.TextMatrix(vsGrid.Row, 1)) = "" Then
        Exit Sub
    End If
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "    " & vsGrid.Row & " 行目をクリアしますか？            "
        .MB_Title = ""
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Sub
    End With
    iB1_Hinbn = ""
    iB1_Seban = ""
    B1lb_Syuyo = ""
    iB1_Maisu = ""
    iB1_Suryo = ""
    B1lb_Hasu = ""
    B1Chk_UMkbn.Value = 0
    Call TextToVsGrid(Val(Input_Area.Tag))
End Sub
Private Sub vsGrid_Scroll()
    'スクロール時はテキストボックスを非表示にします。
'    Call GridTextEnabled(False)
End Sub

'*------------------------------------------*'
'*     取引先名の索引
'*------------------------------------------*'
Private Function Get_Torcd(Torcd As String, TorKB As Integer, Optional DBNo As Integer) As Boolean
    Get_Torcd = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select 略称"
    gSL_Select = gSL_Select & "  from 取引先マスタ "
    gSL_Select = gSL_Select & " where 取引先CD = '" & RTrim(Torcd) & "'"
    Select Case TorKB
        Case 1: gSL_Select = gSL_Select & "   and 取引先区分 = 0 and 契約先区分 = 0"
        Case 2: gSL_Select = gSL_Select & "   and 取引先区分 = 0 and 契約先区分 = 1"
        Case 3: gSL_Select = gSL_Select & "   and 取引先区分 = 0 "
        Case 4: gSL_Select = gSL_Select & "   and 取引先区分 = 1 and 契約先区分 = 0 and 社内区分 = 0"
        Case 5: gSL_Select = gSL_Select & "   and 取引先区分 = 1 and 契約先区分 = 1 and 社内区分 = 0"
        Case 6: gSL_Select = gSL_Select & "   and 取引先区分 = 1 and 契約先区分 = 0 and 社内区分 = 1"
        Case 7: gSL_Select = gSL_Select & "   and 取引先区分 = 1 and 契約先区分 = 0 "
        Case 8: gSL_Select = gSL_Select & "   and 取引先区分 = 1 and 社内区分 = 0"
        Case 9: gSL_Select = gSL_Select & "   and 取引先区分 = 1"
    End Select
    gSL_Select = gSL_Select & " order by 取引先区分 desc"
    If Not TRMRead(gSL_Select, 1, DBNo) Then Exit Function

    TRM_RName = RTrim(TRM.略称)

    Get_Torcd = True
End Function

