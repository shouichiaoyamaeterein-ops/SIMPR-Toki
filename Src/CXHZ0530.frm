VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXHZ0530 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "材料・支給実績入力"
   ClientHeight    =   9945
   ClientLeft      =   1650
   ClientTop       =   1185
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
      TabIndex        =   19
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
         BackColor       =   8438015
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
         BackColor       =   8438015
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
         BackColor       =   8438015
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
      Left            =   8475
      ScaleHeight     =   735
      ScaleWidth      =   6075
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2250
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
      Left            =   195
      ScaleHeight     =   1020
      ScaleWidth      =   6915
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3480
      Width           =   6975
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
            TabIndex        =   5
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
            TabIndex        =   4
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
            TabIndex        =   2
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
            TabIndex        =   3
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
      Height          =   1785
      Left            =   165
      ScaleHeight     =   1725
      ScaleWidth      =   14460
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5205
      Width           =   14520
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   1665
         Index           =   0
         Left            =   30
         Top             =   30
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   2937
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
            Height          =   1605
            Index           =   1
            Left            =   30
            Top             =   30
            Width           =   14340
            _ExtentX        =   25294
            _ExtentY        =   2831
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
               Height          =   1500
               Left            =   30
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   45
               Width           =   14265
               _Version        =   196608
               _ExtentX        =   25162
               _ExtentY        =   2646
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
               FormatString    =   "№ |材　質　　　　　　　　　 |板厚   |幅     |長さ   |背番号  |収容数    |枚数 |端　数    |支給数      |単位|有無償   |"
               Rows            =   4
               Cols            =   16
               BackColor       =   12648384
               BackColorFixed  =   12632256
               BackColorBkg    =   12632256
               FocusRect       =   0
               HighLight       =   0
               ScrollBars      =   0
               RowHeightMin    =   360
               SelectionMode   =   1
               Begin Cis3D_v60.CIS3D Back_Suryo 
                  Height          =   270
                  Left            =   7410
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
               Begin Cis3D_v60.CIS3D Input_Area 
                  Height          =   375
                  Left            =   435
                  Top             =   360
                  Width           =   13635
                  _ExtentX        =   24051
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
                     Left            =   5880
                     TabIndex        =   10
                     Top             =   0
                     Width           =   1080
                     _ExtentX        =   1905
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
                  Begin CisText_V60.CisText iB1_Zaist 
                     Height          =   375
                     Left            =   0
                     TabIndex        =   6
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
                     cDataReplace    =   1
                  End
                  Begin CisText_V60.CisText iB1_Maisu 
                     Height          =   375
                     Left            =   8220
                     TabIndex        =   11
                     Top             =   0
                     Width           =   720
                     _ExtentX        =   1270
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
                     Text            =   "#####"
                     MaxLength       =   5
                     cDataType       =   1
                     cDataReplace    =   1
                     cFbComma        =   0
                     cFdAutoFormat   =   1
                     cGFormat        =   "#####"
                     cILength        =   5
                     cAlignment      =   1
                  End
                  Begin CisText_V60.CisText iB1_Suryo 
                     Height          =   375
                     Left            =   10200
                     TabIndex        =   12
                     Top             =   0
                     Width           =   1575
                     _ExtentX        =   2778
                     _ExtentY        =   661
                     cFormat         =   "#######0.000"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "#######0.000"
                     MaxLength       =   12
                     cDataType       =   2
                     cDataReplace    =   1
                     cFbComma        =   0
                     cFcDecimal      =   3
                     cFdAutoFormat   =   1
                     cGFormat        =   "#######0.000"
                     cILength        =   8
                     cAlignment      =   1
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                     Height          =   360
                     Left            =   6945
                     Top             =   0
                     Width           =   1275
                     _ExtentX        =   2249
                     _ExtentY        =   635
                     BackColor       =   16777152
                     ForeColor       =   0
                     Caption         =   "#####0.000"
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
                     cPositionX      =   -30
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Hasu 
                     Height          =   360
                     Left            =   8925
                     Top             =   0
                     Width           =   1275
                     _ExtentX        =   2249
                     _ExtentY        =   635
                     BackColor       =   16777152
                     ForeColor       =   0
                     Caption         =   "#####0.000"
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
                     cPositionX      =   -30
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Umu 
                     Height          =   360
                     Left            =   12315
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
                        TabIndex        =   13
                        Top             =   45
                        Width           =   825
                     End
                  End
                  Begin CisText_V60.CisText iB1_Itatu 
                     Height          =   375
                     Left            =   3090
                     TabIndex        =   7
                     Top             =   0
                     Width           =   960
                     _ExtentX        =   1693
                     _ExtentY        =   661
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
                     cILength        =   4
                  End
                  Begin CisText_V60.CisText iB1_Width 
                     Height          =   375
                     Left            =   4020
                     TabIndex        =   8
                     Top             =   0
                     Width           =   960
                     _ExtentX        =   1693
                     _ExtentY        =   661
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
                     cILength        =   4
                  End
                  Begin CisText_V60.CisText iB1_Long 
                     Height          =   375
                     Left            =   4950
                     TabIndex        =   9
                     Top             =   0
                     Width           =   960
                     _ExtentX        =   1693
                     _ExtentY        =   661
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
                     cILength        =   4
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Tannm 
                     Height          =   360
                     Left            =   11745
                     Top             =   0
                     Width           =   570
                     _ExtentX        =   1005
                     _ExtentY        =   635
                     BackColor       =   16777152
                     ForeColor       =   0
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
                     cFont3DColor1   =   8421504
                     cFont3DColor2   =   -2147483643
                     cFont3D         =   2
                     cAlingnment     =   7
                     cBoderStyle     =   1
                  End
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
   Begin Cis3D_v60.CIS3D H1_Area1 
      Height          =   1080
      Left            =   180
      Top             =   1995
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1905
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
         Height          =   945
         Left            =   75
         Top             =   75
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1667
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
            Left            =   180
            TabIndex        =   1
            Top             =   450
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
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   714
      BackColor       =   8438015
      ForeColor       =   0
      Caption         =   "【 材料・支給実績入力  】"
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
         BackColor       =   8438015
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
         BackColor       =   8438015
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
Attribute VB_Name = "CXHZ0530"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   材料・支給実績入力
'**       フォームID    :   CXHZ0530
'**       処理概要      :
'**
'**     ---------------------------------------------------
'**
'**       作  成  日    :   2008/08/22  By CIS SIMPR-A
'**       変  更  日    :   2008/09/04  By CIS - 印刷機能を追加
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
    Dim wSuryo          As Currency
    
    Dim Clear_CHT       As Integer      ' 帳票出力テーブル削除用フラグ　0:未削除／1:削除済　2008.08.22 add
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
    SyoriKB = "ADD"
    ProcHB = "H2"               ' 実行開始時は新規入力
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

'   支給初期設定
    wTorKB = 8                          '【発注】手配先/契約先検索
    Back_Ymd.Caption = "納 入 日"
    Back_Torcd1.Caption = "支 給 先"
    Back_Suryo.Caption = "実績数"
    iH2_Bin.Enabled = True
    iH2_Ukeir.Enabled = True
    B1Chk_UMkbn.Enabled = True
'   有無償の初期値を有償にする
    With vsGrid
        For gInt = 1 To .Rows - 1
            .TextMatrix(gInt, 12) = 1
        Next
    End With
    Clear_CHT = 0       ' 帳票出力テーブル削除用フラグクリア（0:未削除）　2008.08.22 add
    
    Call DBDelete_CHT(1)
    
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
'   #### 納品書発行 ####        ' 2008.09.04 add
    Call Prt_Nou                ' 2008.09.04 add
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
      Call HeadBodyClear("B")
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
'        H1Op_JKbn(2).SetFocus      ' 2008.07.09 del
        ProcHB = "H2"               ' 2008.07.09 add
        Call DispChange(ProcHB)     ' 2008.07.09 add
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
        ProcHB = "H1"               ' 2008.07.09 add
        Call DispChange(ProcHB)     ' 2008.07.09 add
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
        ProcHB = "H1"               ' 2008.07.09 add
        Call DispChange(ProcHB)     ' 2008.07.09 add
        iH1_NouNo.SetFocus
        Dummy.Enabled = False
    End If
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
'*--------------------------------------------------<< 2004/06/23 >> Insert Start
    If PB_Look.Tag = "iH1_NouNo" Then Call Look_ZJK: Exit Sub
'*--------------------------------------------------<< 2004/06/23 >> Insert End
    If PB_Look.Tag = "iH2_Torcd1" Then Call Look_Torcd(PB_Look.Tag): Exit Sub
    If PB_Look.Tag = "iH2_Ukeir" Then Call Look_Ukeir: Exit Sub
    If PB_Look.Tag = "iB1_Zaist" Then Call Look_Zaist: Exit Sub
'
    If PB_Look.Tag = "iH2_Ymd" Then iH2_Ymd.ShowCalender: Exit Sub
End Sub
'+----------------------------+
'+       取引先検索            +
'+----------------------------+
Private Sub Look_Torcd(TextName As String)
    RV_Left = 0
    RV_Top = 0

    RV_Call = "C"
    RV_TorKb = wTorKB
    '
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH2_Torcd1 = RV_TorcdK
        H2lb_Tornm1 = RV_TorRName
        Call iH2_Torcd1_GotFocus    ' 2008.07.09 add
        iH2_Torcd1.SetFocus         ' 2008.07.09 add
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
    RV_TorRName = H2lb_Tornm1
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
'*--------------------------------------------------<< 2004/06/23 >> Insert Start

'+----------------------------+
'+      材料検索              +
'+----------------------------+
Private Sub Look_Zaist()
    RV_Left = 0
    RV_Top = 0
    RV_Zaist = ""
    RV_Call = "C"

    CKK0070.Show vbModal
    Unload CKK0070
    Set CKK0070 = Nothing

    If RV_Rtn Then
        iB1_Zaist = RV_Zaist
        iB1_Itatu = RV_Itatu
        iB1_Width = RV_Width
        iB1_Long = RV_Long
        Call iB1_Zaist_LostFocus    ' 2008.09.04 add
    End If
End Sub

'+-------------------------------------+
'+       材料・入出庫入力データ検索
'+-------------------------------------+
Private Sub Look_ZJK()
    RV_Left = 0
    RV_Top = 0
    RV_Call = ""

    CXHZ0531.Show vbModal
    Unload CXHZ0531
    Set CXHZ0531 = Nothing

    If RV_Rtn Then
       iH1_NouNo = RTrim(RV_Code)
    End If
End Sub
'*--------------------------------------------------<< 2004/06/23 >> Insert End

'*--------------------------------------------------<< 2008.09.04 >> Insert Start
'+----------------------------+
'+    印　刷  処　理            +
'+----------------------------+
Private Sub Prt_Nou()
''Exit Sub '試験用
    gSL_Select = "SELECT 検索情報 "
    gSL_Select = gSL_Select & " From 帳票出力テーブル "
    gSL_Select = gSL_Select & " Where 端末番号 =  HOST_NAME() "
    gSL_Select = gSL_Select & "   And 帳票区分 =  302 "
    gSL_Select = gSL_Select & "   And 検索情報 <> '' "
    gSL_Select = gSL_Select & " Group By 検索情報 "
    
    If CHTRead(gSL_Select) Then             ' 帳票出力テーブルに該当データがあれば印刷する
       CXHZ0530_PrtCnt = CisDB.RecordCount  ' 印刷枚数
       Call CXHZ0530LPrint
    End If
    Call CHTClose
End Sub
'*--------------------------------------------------<< 2008.09.04 >> Insert End

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

'---------------------------------------------------------------------- 2008.07.09 add start
'+-----------------------------+
'+      iB1_Zaist
'+-----------------------------+
Private Sub iB1_Zaist_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Set ActObj = iB1_Zaist
    B_TabDummy.Tag = "S"
End Sub
Private Sub iB1_Zaist_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Zaist_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    Call Zaist_Set
End Sub
'+-----------------------------+
'+      iB1_Itatu
'+-----------------------------+
Private Sub iB1_Itatu_GotFocus()
    Set ActObj = iB1_Itatu
End Sub
Private Sub iB1_Itatu_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Itatu_LostFocus()
    Call Zaist_Set
End Sub
'+-----------------------------+
'+      iB1_Width
'+-----------------------------+
Private Sub iB1_Width_GotFocus()
    Set ActObj = iB1_Width
End Sub
Private Sub iB1_Width_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Width_LostFocus()
    Call Zaist_Set
End Sub
'+-----------------------------+
'+      iB1_Long
'+-----------------------------+
Private Sub iB1_Long_GotFocus()
    Set ActObj = iB1_Long
End Sub
Private Sub iB1_Long_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Long_LostFocus()
    Call Zaist_Set
End Sub
'---------------------------------------------------------------------- 2008.07.09 add end

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
                If iB1_Suryo.MaxLength < CisFun.Len2(Format(gLong, "###0.000")) Then    ' "####"
                    SendKeys "{TAB}"
                End If
            End If
    End Select
End Sub
Private Sub iB1_Maisu_LostFocus()
    If CisFun.Val2(iB1_Maisu) <> 0 Then
        If CisFun.Val2(B1lb_Syuyo) <> 0 Then
            gLong = (CisFun.Val2(B1lb_Syuyo) * CisFun.Val2(iB1_Maisu))
            iB1_Suryo = Format(gLong, "###0.000")   ' "####"
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
           'Back_Jkbn.Enabled = True            ' 2008.07.09 del
            Back_NouNo.Enabled = False
            iH1_NouNo.cLostColor = &HFFFFC0
        Case Else
           'Back_Jkbn.Enabled = False           ' 2008.07.09 del
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
    PB_ADD.Visible = H1Mode
    PB_MNT.Visible = False: If Not B1Mode And SyoriKB <> "MNT" Then PB_MNT.Visible = True
    PB_DEL.Visible = False: If (H1Mode And SyoriKB <> "DEL") Or (H2Mode And SyoriKB = "ADD") Then PB_DEL.Visible = True
    PB_CAN.Visible = False: If B1Mode Or (H2Mode And SyoriKB <> "ADD") Then PB_CAN.Visible = True
    PB_END.Visible = False: If H1Mode Or (H2Mode And SyoriKB = "ADD") Then PB_END.Visible = True

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
        '*------------------------- 2008.07.09 add start
        '有無償の初期値をセットする (1：有償)
        With vsGrid
            For gInt = 1 To .Rows - 1
                .TextMatrix(gInt, 12) = 1
            Next
        End With
        '*------------------------- 2008.07.09 add end
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
    gSL_Select = gSL_Select & "  zjk.*"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  材料実績管理テーブル zjk"
    gSL_Select = gSL_Select & " where zjk.納入番号 = '" & RTrim(iH1_NouNo) & "'"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  zjk.納入番号"
    gSL_Select = gSL_Select & ", zjk.納品書行"
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
'        If H1Op_JKbn(3) Then
            CisFun.MB_MSG(2) = "    納入便を入力して下さい            "
            CisFun.MBOX
            iH2_Bin.SetFocus
            Exit Function
'        End If
    End If
'   入庫先 or 出庫先 or 支給先
    If Trim(iH2_Torcd1) = "" Then
'        If H1Op_JKbn(3) Then
            CisFun.MB_MSG(2) = "    " & Back_Torcd1.Caption & "を入力して下さい            "            ' 支給先
            CisFun.MBOX
            iH2_Torcd1.SetFocus
            Exit Function
'        End If
    Else
''''''''        H2lb_Tornm1 = ""
        If Not Get_Torcd(iH2_Torcd1, wTorKB) Then
            CisFun.MB_MSG(2) = "    " & Back_Torcd1.Caption & "　未登録！（取引先マスタ）            "  ' 支給先
            CisFun.MBOX
            iH2_Torcd1.SetFocus
            Exit Function
        End If
    End If
'   受入
    If RTrim(iH2_Ukeir) = "" Then
'        If H1Op_JKbn(3) Then
            CisFun.MB_MSG(2) = "    受入を入力して下さい            "
            CisFun.MBOX
            iH2_Ukeir.SetFocus
            Exit Function
'        End If
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

    With vsGrid
        .Redraw = False
        For gInt = 1 To .Rows - 1   ' 2008.07.09 行数変更　10→3：グリッドの行数-1
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
                        For wColNo1 = 1 To 13
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
            Set ErrObj = iB1_Zaist
            wRowNo1 = 1
            GoTo Body1Chk_Err
        End If
'
        For wRowNo1 = 1 To .Rows - 1
            If RTrim(.TextMatrix(wRowNo1, 1)) = "" Then Exit For

        '   << 重複チェック >>
            For wRowNo2 = wRowNo1 + 1 To .Rows - 1
                If RTrim(.TextMatrix(wRowNo1, 1)) = RTrim(.TextMatrix(wRowNo2, 1)) And _
                   CisFun.Val2(.TextMatrix(wRowNo1, 2)) = CisFun.Val2(.TextMatrix(wRowNo2, 2)) And _
                   CisFun.Val2(.TextMatrix(wRowNo1, 3)) = CisFun.Val2(.TextMatrix(wRowNo2, 3)) And _
                   CisFun.Val2(.TextMatrix(wRowNo1, 4)) = CisFun.Val2(.TextMatrix(wRowNo2, 4)) Then
                    CisFun.MB_MSG(1) = "    材料が重複しています。            "
                    CisFun.MB_MSG(3) = "    " & wRowNo1 & "行目と " & wRowNo2 & "行目            "
                    CisFun.MBOX
                    Set ErrObj = iB1_Zaist
                    wRowNo1 = wRowNo2
                    GoTo Body1Chk_Err
                End If
            Next wRowNo2
        ' 背番号半角チェック
            If RTrim(.TextMatrix(wRowNo1, 5)) <> "" And Not HalfSizeChk(.TextMatrix(wRowNo1, 5)) Then
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
        '   << 材料マスタ存在チェック >>
            gSL_Select = ""
            gSL_Select = gSL_Select & "select *"
            gSL_Select = gSL_Select & "  from 材料マスタ"
            gSL_Select = gSL_Select & " where 材質 = '" & RTrim(.TextMatrix(wRowNo1, 1)) & "'"
            gSL_Select = gSL_Select & "   and 板厚 = " & CisFun.Val2(.TextMatrix(wRowNo1, 2)) & ""
            gSL_Select = gSL_Select & "   and 幅 = " & CisFun.Val2(.TextMatrix(wRowNo1, 3)) & ""
            gSL_Select = gSL_Select & "   and 長さ = " & CisFun.Val2(.TextMatrix(wRowNo1, 4)) & ""
'Call Test_LogOutput(gSL_Select, 1)  '<<<<<< Debug
            If Not ZRMRead(gSL_Select, 1) Then
                CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　材質・板厚・幅・長さ　未登録！（材料マスタ）            "
                CisFun.MBOX
                Set ErrObj = iB1_Zaist
                GoTo Body1Chk_Err
            End If
        
        '*------------------------------------------------- 2008.07.10 add start
        '   << 材料取引先マスタ存在チェック >>
            'ZTM.材質 = RTrim(.TextMatrix(wRowNo1, 1))
'            ZTM.材料管理番号 = ZRM.材料管理番号            '材料マスタの材料管理番号から検索する
            gSL_Select = ""
            gSL_Select = gSL_Select & "select *"
            gSL_Select = gSL_Select & "  from 材料取引先マスタ"
            gSL_Select = gSL_Select & " where 材質 = '" & RTrim(.TextMatrix(wRowNo1, 1)) & "'"
            gSL_Select = gSL_Select & "   and 板厚 = " & CisFun.Val2(.TextMatrix(wRowNo1, 2)) & ""
            gSL_Select = gSL_Select & "   and 幅 = " & CisFun.Val2(.TextMatrix(wRowNo1, 3)) & ""
            gSL_Select = gSL_Select & "   and 長さ = " & CisFun.Val2(.TextMatrix(wRowNo1, 4)) & ""
            gSL_Select = gSL_Select & "   and 取引先 = '" & RTrim(iH2_Torcd1) & "' "    ' 2008.09.04 add
            gSL_Select = gSL_Select & "   and 受入   = '" & RTrim(iH2_Ukeir) & "' "     ' 2008.09.05 add
            'gSL_Select = gSL_Select & "   and 品目   = 1 "  ' 発注                      ' 2008.09.04 add
'Call Test_LogOutput(gSL_Select, 1)  '<<<<<< Debug
'            If Not ZTMRead("", 1) Then
            If Not ZTMRead(gSL_Select, 1) Then
                CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　材質　未登録！（材料取引先マスタ）            "
                CisFun.MBOX
                Set ErrObj = iB1_Zaist
                GoTo Body1Chk_Err
            End If
        '   << 材料取引先マスタ　手配区分チェック >>
            If ZTM.手配区分 <> "4" Then
                CisFun.MB_MSG(2) = "    【" & wRowNo1 & "行目】　材料の手配区分が支給ではありません  （材料取引先マスタ）            "
                CisFun.MBOX
                Set ErrObj = iB1_Zaist
                GoTo Body1Chk_Err
            End If
        '*------------------------------------------------- 2008.07.10 add end
               
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
            If CisFun.Val2(.TextMatrix(wRowNo1, 9)) = 0 Then
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
        With ZJK
            iH1_NouNo = RTrim(.納入番号)
            iH2_Ymd = .実績日
            iH2_Bin = .実績便
            iH2_Torcd1 = RTrim(.実績先)
            iH2_Ukeir = RTrim(.受入)
            If Get_Torcd(.実績先, 0) Then H2lb_Tornm1 = RTrim(TRM.略称)
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
    Call DBDelete_ZJK
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
    
'' ( 削除処理 )
''---<< 2004.12.30 追加更新時の削除はなし  >>---------------------------Start------------
    If SyoriKB = "MNT" Then
        Call DBDelete_ZJK
        Call DBDelete_CHT(0)        ' 2008.09.04 add
    End If
''---<< 2004.12.30 追加更新時の削除はなし  >>---------------------------End  ------------

' ( 削除処理 )
'    Call DBDelete_ZJK              ' 2008.09.04 del
' ( 追加処理 )
    If Not DBSet_ZJK Then GoTo DBPut_Err
    For gIndex = 1 To vsGrid.Rows - 1
        If RTrim(vsGrid.TextMatrix(gIndex, 1)) <> "" Then
            If DBInsert_ZJK(gIndex) Then    ' Call DBInsert_ZJK(gIndex)     2008.09.04 upd
                If SyoriKB = "ADD" And _
                  Clear_CHT = 0 Then        ' 自端末起動後最初の登録ならば  2008.09.04 add
                    Call DBDelete_CHT(1)    ' 帳票出力テーブルの            2008.09.04 add
                    Clear_CHT = 1           ' 該当データを削除する          2008.09.04 add
                End If                      '                               2008.09.04 add
                Call DBInsert_CHT(gIndex)   ' 帳票出力テーブル登録          2008.09.04 add
            End If                          '                               2008.09.04 add
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
'+      削除処理 - 材料実績管理テーブル
'+------------------------------------+
Private Function DBDelete_ZJK(Optional DBNo As Integer) As Boolean
    DBDelete_ZJK = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "delete 材料実績管理テーブル"
    gSL_Select = gSL_Select & " where 納入番号 = '" & RTrim(iH1_NouNo) & "'"
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    Call DBDelete_CHT(0)    ' 印刷用データを削除
    
    DBDelete_ZJK = True
End Function
'*------------------------------------------------------------------------------ 2008.09.04 add start
'+------------------------------------+
'+      削除処理 - 帳票出力テーブル
'+          Mode = 0:自端末の該当納入番号のみ
'+                 1:自端末の納品書発行対象データ全て
'+------------------------------------+
Private Function DBDelete_CHT(Mode As Integer, Optional DBNo As Integer) As Boolean
    DBDelete_CHT = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "delete 帳票出力テーブル"
    gSL_Select = gSL_Select & " where 端末番号 = substring(host_name(),1,20) "
    gSL_Select = gSL_Select & "   And 帳票区分 = 302 "
    If Mode = 0 Then
        gSL_Select = gSL_Select & "   And 検索情報 = '" & RTrim(iH1_NouNo) & "'"
    End If

    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With

    DBDelete_CHT = True
End Function
'*------------------------------------------------------------------------------ 2008.09.04 add end

'+------------------------------------+
'+      セット - 材料実績管理テーブル
'+------------------------------------+
Private Function DBSet_ZJK()
    DBSet_ZJK = False

    Call ItemsClearZJK
    With ZJK
        .実績区分 = 3                       ' 2008.07.09 add
        .実績先 = RTrim(iH2_Torcd1)
        .受入 = RTrim(iH2_Ukeir)
        .実績日 = RTrim(iH2_Ymd)
        .実績便 = RTrim(iH2_Bin)
        .作成者 = gTanto
    End With

    If SyoriKB = "ADD" Then
'       ## 追加の場合 ##
        With CisDB
        '   <><><> 納入番号の獲得 <><><>
            .SQL = "材料実績納入番号獲得"
            .StoadoCount = 2
            .ParaNo = 0: .ParaIO = Return用
            .ParaNo = 1: .ParaIO = OutPut用
            .ParaNo = 2: .ParaIO = OutPut用
        '
            If Not .DBStored Then
                CisFun.MB_Lines = 4
                CisFun.MB_MSG(2) = "    材料実績納入番号の獲得に失敗しました            "
                CisFun.MB_Title = "採番処理"
                CisFun.MB_Button = Error
                CisFun.MBOX
                Exit Function
            Else
                .ParaNo = 1: ZJK.納入番号 = .ParaValue & "0"
            End If
        End With
    Else
'       ## 修正の場合 ##
        ZJK.納入番号 = RTrim(iH1_NouNo)
    End If

    DBSet_ZJK = True
End Function

'+------------------------------------+
'+      追加処理 - 実績管理テーブル
'+------------------------------------+
Private Function DBInsert_ZJK(ID As Long, Optional DBNo As Integer) As Boolean
    DBInsert_ZJK = False

'*--------------------------------------------------<< 2004/06/23 >> Update Start
'   If SyoriKB = "ADD" Then
    If CisFun.Val2(vsGrid.TextMatrix(ID, 13)) = 0 Then
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
            .ParaNo = 1: .ParaValue = 10
        '
            If Not .DBStored Then
                CisFun.MB_Lines = 4
                CisFun.MB_MSG(2) = "    材料実績管理番号の獲得に失敗しました            "
                CisFun.MB_Title = "採番処理"
                CisFun.MB_Button = Error
                CisFun.MBOX
                Exit Function
            Else
                .ParaNo = 2: ZJK.実績管理NO = .ParaValue
            End If
        End With
    Else
'       ## 修正の場合 ##
        ZJK.実績管理NO = CisFun.Val2(vsGrid.TextMatrix(ID, 13))
    End If

    With ZJK
        .納品書行 = ID
        .材質 = RTrim(vsGrid.TextMatrix(ID, 1))
        .板厚 = CisFun.Val2(vsGrid.TextMatrix(ID, 2))
        .幅 = CisFun.Val2(vsGrid.TextMatrix(ID, 3))
        .長さ = CisFun.Val2(vsGrid.TextMatrix(ID, 4))
        .背番号 = RTrim(vsGrid.TextMatrix(ID, 5))
        .収容数 = CisFun.Val2(vsGrid.TextMatrix(ID, 6))
        .実績枚数 = CisFun.Val2(vsGrid.TextMatrix(ID, 7))
        .実績端数 = CisFun.Val2(vsGrid.TextMatrix(ID, 8))
        .実績数 = CisFun.Val2(vsGrid.TextMatrix(ID, 9))
        .有無償区分 = CisFun.Val2(vsGrid.TextMatrix(ID, 12))
    End With
    Call ZJKInsert

    DBInsert_ZJK = True
End Function
'
'*---------------------------------------------------------- 2008.09.04 add start
'+------------------------------------+
'+      追加処理 - 帳票出力テーブル
'+------------------------------------+
Private Function DBInsert_CHT(ID As Long) As Boolean
    DBInsert_CHT = False

    gSL_Select = "Select 納品書発行区分 "
    gSL_Select = gSL_Select & " From  材料取引先マスタ "
    gSL_Select = gSL_Select & " Where 材質   = '" & ZJK.材質 & "' "
    gSL_Select = gSL_Select & "   And 板厚   = " & ZJK.板厚
    gSL_Select = gSL_Select & "   And 幅   = " & ZJK.幅
    gSL_Select = gSL_Select & "   And 長さ   = " & ZJK.長さ
    gSL_Select = gSL_Select & "   And 取引先 = '" & ZJK.実績先 & "' "
    gSL_Select = gSL_Select & "   And 受入   = '" & ZJK.受入 & "' "
'Call Test_LogOutput(gSL_Select, 1)
    If Not ZTMRead(gSL_Select, 1, 1) Then GoTo DBInsert_CHT_Err
    
    If ZTM.納品書発行区分 = 0 Then
       Call ItemsClearCHT
       With CHT
            .端末番号 = "' + substring(host_name(),1,20) + '"
            .帳票区分 = 302
           '.管理文字 = ""
            .検索情報 = ZJK.納入番号
            .管理NO = ZJK.実績管理NO
            .管理行NO = 0
           '.補助番号 = 0
           '.出力先 = 0
        End With
    Call CHTInsert
    End If
    
    DBInsert_CHT = True
DBInsert_CHT_Err:
End Function
'*---------------------------------------------------------- 2008.09.04 add end

'*------------------------------------------*'
'*    材料による背番号・収容数の表示          *
'*------------------------------------------*'
Private Sub Zaist_Set()
    If RTrim(iB1_Zaist) = "" Then Exit Sub

    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  isnull(sym.値名称, '') 手配単位名"
    gSL_Select = gSL_Select & ", isnull(ztm.背番号, '') 背番号"
    gSL_Select = gSL_Select & ", isnull(ztm.変更日, '') 変更日"
    gSL_Select = gSL_Select & ", isnull(ztm.収容数１, 0) 収容数１"
    gSL_Select = gSL_Select & ", isnull(ztm.収容数２, 0) 収容数２"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  材料マスタ zrm"
    gSL_Select = gSL_Select & "  left outer join 材料取引先マスタ ztm"
    gSL_Select = gSL_Select & "    on ztm.材料管理番号 = zrm.材料管理番号"
    gSL_Select = gSL_Select & "   and ztm.取引先 = '" & RTrim(iH2_Torcd1) & "'"
'*--------------------------------------------------<< 2004/06/23 >> Update Start
'   gSL_Select = gSL_Select & "   and ztm.受入 = '" & RTrim(iH2_Ukeir) & "'"
    If RTrim(iH2_Ukeir) <> "" Then
        gSL_Select = gSL_Select & "   and ztm.受入 = '" & RTrim(iH2_Ukeir) & "'"
    End If
'*--------------------------------------------------<< 2004/06/23 >> Update End
    gSL_Select = gSL_Select & "  left outer join 名称マスタ sym"
    gSL_Select = gSL_Select & "    on sym.区分名称 = '単位区分'"
    gSL_Select = gSL_Select & "   and sym.値 = zrm.単位"
    gSL_Select = gSL_Select & "   and sym.区分タイプ = ''"
    gSL_Select = gSL_Select & " where zrm.材質 = '" & RTrim(iB1_Zaist) & "'"
    gSL_Select = gSL_Select & "   and zrm.板厚 = " & CisFun.Val2(iB1_Itatu) & ""
    gSL_Select = gSL_Select & "   and zrm.幅 = " & CisFun.Val2(iB1_Width) & ""
    gSL_Select = gSL_Select & "   and zrm.長さ = " & CisFun.Val2(iB1_Long) & ""
'*--------------------------------------------------<< 2004/06/23 >> Insert Start
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  isnull(ztm.受入, '')"
'*--------------------------------------------------<< 2004/06/23 >> Insert End
    If ZTMRead(gSL_Select, 1) Then
        iB1_Seban = RTrim(ZTM.背番号)
        B1lb_Tannm = RTrim(ZTM.手配単位名)
'
        Call Syuyo_Set      ' 収容数表示
        Call Suryo_Set      ' 枚数・端数・実績数の再表示
'*--------------------------------------------------<< 2004/06/23 >> Insert Start
    Else
       'iB1_Seban = ""
        B1lb_Tannm = ""
        B1lb_Syuyo = ""
        Call Suryo_Set      ' 枚数・端数・実績数の再表示
'*--------------------------------------------------<< 2004/06/23 >> Insert End
    End If
End Sub

'*------------------------------------------*'
'*    収容数の表示                           *
'*------------------------------------------*'
Private Sub Syuyo_Set()
    If RTrim(ZTM.変更日) <> "" Then
        If RTrim(iH2_Bin) <> "" Then
            If RTrim(ZTM.変更日) & RTrim(ZTM.変更便) > RTrim(iH2_Ymd) & RTrim(iH2_Bin) Then
                B1lb_Syuyo = Format(ZTM.収容数1, "###0.000")
            Else
                B1lb_Syuyo = Format(ZTM.収容数2, "###0.000")
            End If
        Else
            If RTrim(ZTM.変更日) > RTrim(iH2_Ymd) Then
                B1lb_Syuyo = Format(ZTM.収容数1, "###0.000")
            Else
                B1lb_Syuyo = Format(ZTM.収容数2, "###0.000")
            End If
        End If
    Else
        B1lb_Syuyo = Format(ZTM.収容数1, "###0.000")
    End If
End Sub

'*------------------------------------------*'
'*    枚数・端数・実績数の再表示              *
'*------------------------------------------*'
Private Sub Suryo_Set()
    Dim wMSuu       As Currency
    Dim wHSuu       As Currency

    B1lb_Hasu = ""
    If CisFun.Val2(B1lb_Syuyo) <> 0 Then
        If RTrim(iB1_Suryo) = "" Then
            If CisFun.Val2(iB1_Maisu) <> 0 Then
                wSuryo = (CisFun.Val2(B1lb_Syuyo) * CisFun.Val2(iB1_Maisu))
                iB1_Suryo = Format(wSuryo, "###0.000")
            End If
        Else
            wMSuu = Round(CisFun.Val2(iB1_Suryo) / CisFun.Val2(B1lb_Syuyo))
            wHSuu = CisFun.Val2(iB1_Suryo) - (CisFun.Val2(B1lb_Syuyo) * wMSuu)
            If wHSuu > 0 Then wMSuu = wMSuu + 1
            If wHSuu < 0 Then wHSuu = wHSuu + CisFun.Val2(B1lb_Syuyo)
            iB1_Maisu = Format(wMSuu, "####")
            If wHSuu <> 0 Then
                B1lb_Hasu = Format(wHSuu, "###0.000")
            End If
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
        Call .FixedSet(中中, 左中, あり, 25, 1, "材　質")
        Call .FixedSet(中中, 右中, あり, 7, 2, "板厚")
        Call .FixedSet(中中, 右中, あり, 7, 3, "幅")
        Call .FixedSet(中中, 右中, あり, 7, 4, "長さ")
        Call .FixedSet(中中, 左中, あり, 8, 5, "背番号")
        Call .FixedSet(中中, 右中, あり, 10, 6, "収容数")
        Call .FixedSet(中中, 右中, あり, 5, 7, "枚数")
        Call .FixedSet(中中, 右中, あり, 10, 8, "端　数")
        Call .FixedSet(中中, 右中, あり, 12, 9, "実績数")
        Call .FixedSet(中中, 中中, あり, 4, 10, "単位")
        Call .FixedSet(中中, 中中, なし, 9, 11, "有無償")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 12, "有無償区分")
        Call .FixedSet(中中, 左中, なし, 0, 13, "実績管理NO")
        Call .FixedSet(中中, 左中, なし, 0, 14, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11")
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

    If ZJKRead(gSL_Select, , 1) Then
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
        Do Until Not ZJK_RDSTS
            ID = ID + 1

            Call ZjkInfoToGrid(ID)

            Call ZJKReadNext(1)
        Loop
        With vsGrid
             .Row = 1: .Col = 1: .ColSel = .Cols - 1
             .Redraw = True
        End With

        GridSet = True
    End If
    Call ZJKClose
End Function

'+----------------------------------------------+
'+      材料実績管理テーブルよりグリッドに転送
'+----------------------------------------------+
Private Sub ZjkInfoToGrid(SetRow As Long)

    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  sym.値名称"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  材料マスタ zrm"
    gSL_Select = gSL_Select & "  left outer join 名称マスタ sym"
    gSL_Select = gSL_Select & "    on sym.区分名称 = '単位区分'"
    gSL_Select = gSL_Select & "   and sym.値 = zrm.単位"
    gSL_Select = gSL_Select & "   and sym.区分タイプ = ''"
    gSL_Select = gSL_Select & " where zrm.材質 = '" & RTrim(ZJK.材質) & "'"
    gSL_Select = gSL_Select & "   and zrm.板厚 = " & ZJK.板厚 & ""
    gSL_Select = gSL_Select & "   and zrm.幅 = " & ZJK.幅 & ""
    gSL_Select = gSL_Select & "   and zrm.長さ = " & ZJK.長さ & ""
    If Not SYMRead(gSL_Select, 1) Then
        SYM.値名称 = ""
    End If
'
    With vsGrid
        'No
        .TextMatrix(SetRow, 0) = SetRow
        '材質
        .TextMatrix(SetRow, 1) = RTrim(ZJK.材質)
        '板厚
        If ZJK.板厚 <> 0 Then
            .TextMatrix(SetRow, 2) = Format(ZJK.板厚, "###0.00")
        End If
        '幅
        If ZJK.幅 <> 0 Then
            .TextMatrix(SetRow, 3) = Format(ZJK.幅, "###0.00")
        End If
        '長さ
        If ZJK.長さ <> 0 Then
            .TextMatrix(SetRow, 4) = Format(ZJK.長さ, "###0.00")
        End If
        '背番号
        .TextMatrix(SetRow, 5) = RTrim(ZJK.背番号)
        '収容数
        If ZJK.収容数 <> 0 Then
            .TextMatrix(SetRow, 6) = Format(ZJK.収容数, "###0.000")
        End If
        '枚数
        .TextMatrix(SetRow, 7) = Format(ZJK.実績枚数, "####")
        '端数
        If ZJK.実績端数 <> 0 Then
            .TextMatrix(SetRow, 8) = Format(ZJK.実績端数, "###0.000")
        End If
        '実績数
        If ZJK.実績数 <> 0 Then
            .TextMatrix(SetRow, 9) = Format(ZJK.実績数, "###0.000")
        End If
        '単位
        .TextMatrix(SetRow, 10) = RTrim(SYM.値名称)
        '有無償
        If ZJK.実績区分 = 3 Then
            Select Case ZJK.有無償区分
                Case 0: .TextMatrix(SetRow, 11) = "無償"
                Case 1: .TextMatrix(SetRow, 11) = "有償"
            End Select
        End If
        .TextMatrix(SetRow, 12) = ZJK.有無償区分
        .TextMatrix(SetRow, 13) = ZJK.実績管理NO
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
                If vsGrid.Row < vsGrid.Rows - 1 Then            ' 2008.07.09 add : 行Maxを超えないようにする
                    vsGrid.Rows = gInt + 1 + 1
                    vsGrid.TextMatrix(gInt + 1, 0) = gInt + 1
                Else                                            ' 2008.07.09 add : 行Maxを超えないようにする
                    GoTo Grid_RowChange_End                     ' 2008.07.09 add : 行Maxを超えないようにする
                End If
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
        '材質
        .TextMatrix(RowNo, 1) = RTrim(iB1_Zaist)
        '板厚
        If CisFun.Val2(iB1_Itatu) <> 0 Then
            .TextMatrix(RowNo, 2) = Format(CisFun.Val2(iB1_Itatu), "###0.00")
        Else
            .TextMatrix(RowNo, 2) = ""
        End If
        '幅
        If CisFun.Val2(iB1_Width) <> 0 Then
            .TextMatrix(RowNo, 3) = Format(CisFun.Val2(iB1_Width), "###0.00")
        Else
            .TextMatrix(RowNo, 3) = ""
        End If
        '長さ
        If CisFun.Val2(iB1_Long) <> 0 Then
            .TextMatrix(RowNo, 4) = Format(CisFun.Val2(iB1_Long), "###0.00")
        Else
            .TextMatrix(RowNo, 4) = ""
        End If
        '背番号
        .TextMatrix(RowNo, 5) = RTrim(iB1_Seban)
        '収容数
        .TextMatrix(RowNo, 6) = B1lb_Syuyo
        '枚数
        If CisFun.Val2(iB1_Maisu) <> 0 Then
            .TextMatrix(RowNo, 7) = Format(CisFun.Val2(iB1_Maisu), "####")
        Else
            .TextMatrix(RowNo, 7) = ""
        End If
        '端数
        .TextMatrix(RowNo, 8) = B1lb_Hasu
        '実績数
        If CisFun.Val2(iB1_Suryo) <> 0 Then
            .TextMatrix(RowNo, 9) = Format(CisFun.Val2(iB1_Suryo), "###0.000")
        Else
            .TextMatrix(RowNo, 9) = ""
        End If
        '単位
        .TextMatrix(RowNo, 10) = RTrim(B1lb_Tannm)
        '有無償
        .TextMatrix(RowNo, 11) = ""
        If RTrim(iB1_Zaist) <> "" Then
            .TextMatrix(RowNo, 11) = B1Chk_UMkbn.Caption        ' If H1Op_JKbn(3) Then
        End If
        .TextMatrix(RowNo, 12) = B1Chk_UMkbn.Value
    End With
End Sub

'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(RowNo As Integer)
    With vsGrid
        Input_Area.Tag = RowNo
        Input_Area.Move Input_Area.Left, vsGrid.CellTop, Input_Area.Width, vsGrid.CellHeight
        '材質
        iB1_Zaist = RTrim(.TextMatrix(RowNo, 1))
        '板厚
        iB1_Itatu = .TextMatrix(RowNo, 2)
        '幅
        iB1_Width = .TextMatrix(RowNo, 3)
        '長さ
        iB1_Long = .TextMatrix(RowNo, 4)
        '背番号
        iB1_Seban = RTrim(.TextMatrix(RowNo, 5))
        '収容数
        B1lb_Syuyo = .TextMatrix(RowNo, 6)
        '枚数
        If CisFun.Val2(.TextMatrix(RowNo, 7)) <> 0 Then
            iB1_Maisu = CisFun.Val2(.TextMatrix(RowNo, 7))
        Else
            iB1_Maisu = ""
        End If
        '端数
        B1lb_Hasu = .TextMatrix(RowNo, 8)
        '支給数
        If CisFun.Val2(.TextMatrix(RowNo, 9)) <> 0 Then
            iB1_Suryo = CisFun.Val2(.TextMatrix(RowNo, 9))
        Else
            iB1_Suryo = ""
        End If
        '単位
        B1lb_Tannm = RTrim(.TextMatrix(RowNo, 10))
        '有無償
        B1Chk_UMkbn = CisFun.Val2(.TextMatrix(RowNo, 12))
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
    iB1_Zaist = ""
    iB1_Itatu = ""
    iB1_Width = ""
    iB1_Long = ""
    iB1_Seban = ""
    B1lb_Syuyo = ""
    iB1_Maisu = ""
    B1lb_Hasu = ""
    iB1_Suryo = ""
    B1lb_Tannm = ""
    B1Chk_UMkbn.Value = 1   ' 初期値は有償
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
