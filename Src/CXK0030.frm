VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXK0030 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "検収画面入力 [単品]"
   ClientHeight    =   9945
   ClientLeft      =   90
   ClientTop       =   1710
   ClientWidth     =   15375
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
   ScaleWidth      =   15375
   WindowState     =   2  '最大化
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'なし
      Height          =   345
      Index           =   1
      Left            =   10260
      ScaleHeight     =   345
      ScaleWidth      =   3675
      TabIndex        =   14
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
      Height          =   1320
      Left            =   9720
      ScaleHeight     =   1260
      ScaleWidth      =   3435
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1440
      Width           =   3495
      Begin CisBtn_60.CisBtn PB_CAN 
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
      Left            =   1830
      ScaleHeight     =   870
      ScaleWidth      =   6630
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1830
      Width           =   6690
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   780
         Left            =   3720
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
            TabIndex        =   1
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
         Caption         =   "納 入 先"
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
         Left            =   4425
         Top             =   60
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "検 収 日"
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_Ymd 
            Height          =   345
            Left            =   105
            TabIndex        =   2
            Top             =   330
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
      TabIndex        =   11
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
      Height          =   6240
      Left            =   1845
      ScaleHeight     =   6180
      ScaleWidth      =   11700
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3420
      Width           =   11760
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   6075
         Index           =   0
         Left            =   60
         Top             =   60
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   10716
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
            Height          =   6015
            Index           =   1
            Left            =   30
            Top             =   30
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   10610
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
               Height          =   5850
               Left            =   90
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   75
               Width           =   11370
               _Version        =   196608
               _ExtentX        =   20055
               _ExtentY        =   10319
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
               FormatString    =   "№ |納入番号  | 背番号  |品　番　　　　　　　　　 |赤/黒|収容数  |枚　数 |端　数 |入荷数 　||"
               Rows            =   16
               Cols            =   20
               BackColor       =   12648384
               BackColorBkg    =   12632256
               FocusRect       =   0
               HighLight       =   0
               RowHeightMin    =   360
               SelectionMode   =   1
               Begin Cis3D_v60.CIS3D Input_Area 
                  Height          =   345
                  Left            =   450
                  Top             =   360
                  Width           =   10425
                  _ExtentX        =   18389
                  _ExtentY        =   609
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
                  Begin Cis3D_v60.CIS3D B1_Back 
                     Height          =   360
                     Left            =   15
                     Top             =   0
                     Width           =   6240
                     _ExtentX        =   11007
                     _ExtentY        =   635
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
                     cPositionY      =   30
                     Begin VB.CheckBox B1Chk_RBKbn 
                        BackColor       =   &H00C0C0C0&
                        Caption         =   "黒"
                        ForeColor       =   &H00000000&
                        Height          =   225
                        Left            =   5625
                        TabIndex        =   6
                        Top             =   60
                        Width           =   585
                     End
                     Begin CisText_V60.CisText iB1_Seban 
                        Height          =   375
                        Left            =   1275
                        TabIndex        =   4
                        Top             =   0
                        Width           =   1155
                        _ExtentX        =   2037
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
                        Left            =   2415
                        TabIndex        =   5
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
                     Begin CisText_V60.CisText iB1_Nouno 
                        Height          =   375
                        Left            =   0
                        TabIndex        =   3
                        Top             =   0
                        Width           =   1305
                        _ExtentX        =   2302
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
                        Text            =   "XXXXXXXXX"
                        MaxLength       =   9
                        IMEMode         =   2
                        cDataReplace    =   1
                     End
                  End
                  Begin CisText_V60.CisText iB1_Maisu 
                     Height          =   375
                     Left            =   7275
                     TabIndex        =   7
                     Top             =   0
                     Width           =   945
                     _ExtentX        =   1667
                     _ExtentY        =   661
                     cFormat         =   "###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "###,###"
                     MaxLength       =   7
                     cDataType       =   1
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "######"
                     cILength        =   6
                  End
                  Begin CisText_V60.CisText iB1_HSuryo 
                     Height          =   360
                     Left            =   9120
                     TabIndex        =   8
                     Top             =   0
                     Width           =   1200
                     _ExtentX        =   2117
                     _ExtentY        =   635
                     cFormat         =   "#,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "#,###,###"
                     MaxLength       =   9
                     cDataType       =   1
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "#######"
                     cILength        =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                     Height          =   360
                     Left            =   6240
                     Top             =   0
                     Width           =   1035
                     _ExtentX        =   1826
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
                     Left            =   8205
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
      Caption         =   "【 検収画面入力 [単品]  】"
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
End
Attribute VB_Name = "CXK0030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   検収画面入力 [単品]
'**       フォームID    :   CXK0030
'**       処理概要      :
'**
'**       作  成  日    :   2004/03/27  By CIS SIMPR-A
'**       変  更  日    :   2004/04/15  納入先検索・納入先名表示時の取引先区分=3に修正
'**                                     品番取引先マスタの参照表示に品目=0追加
'**       変  更  日    :   2004/05/18  By CIS - 入力処理時、グリッド移動による不具合修正（クリック時）
'**       変  更  日    :   2004/05/28  By CIS 背番号半角チェック追加
'**       変  更  日    :   2005/12/14  By CIS 0除算修正
'**       変  更  日    :   2006/06/06  By CIS 明細行数拡張チェック変更（背番→品番）
'**       変  更  日    :   2008/04/10  By CIS 背番号8桁
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKb         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim ActObj          As Object
    Dim ErrObj          As Object
    Dim ActRow          As Integer
    Dim BinMax          As Byte         '設定最大便（８便）
    Dim mCycle          As String       'サイクル退避
    Dim wsURYM          As String
    
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
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = vsGrid
        .NoSpaceCol = 1
        .NumberSet = All
        .FixedHeight = 360
        .PatanMax = 1
        .SelectCol = 18
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
    SyoriKb = "ADD"
    Call SyoriName(SyoriKb)
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
    
    iH1_Ymd = Format(Now(), "yyyymmdd")
' 売上実績年月
    wsURYM = ""
    Call ItemsClearSYM
    gSL_Select = "Select * From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = '処理内容'"
    gSL_Select = gSL_Select & " And   区分タイプ = ''"
    gSL_Select = gSL_Select & " And   値 = 'PXU0110'"
    If SYMRead(gSL_Select, 1) Then
        With SYM
            wsURYM = SYM.英数字1
        End With
    End If
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
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_Can_Click    '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      '【入力】
                                 'Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
                                 'Dummy.Enabled = False
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
'+   終了処理
'+---------------------+
'    Call EndProc
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
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call ReturnPress
    Dummy.Enabled = False
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_Can_Click()
   Dummy.Enabled = True: Dummy.SetFocus: DoEvents
   Call CisFrm.MousePT(11)
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKb)
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
    If SyoriKb <> "ADD" Then
        SyoriKb = "ADD"
        Call SyoriName(SyoriKb)
    End If
    iH1_Torcd.SetFocus
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If SyoriKb <> "MNT" Then
        SyoriKb = "MNT"
        Call SyoriName(SyoriKb)
    End If
    iH1_Torcd.SetFocus
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKb <> "DEL" Then
        SyoriKb = "DEL"
    Else
        SyoriKb = ""
    End If
    Call SyoriName(SyoriKb)
    iH1_Torcd.SetFocus
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_Ukeir: Exit Sub
    If PB_Look.Tag = "iH1_Ymd" Then
       iH1_Ymd.ShowCalender
       Exit Sub
    End If
    If PB_Look.Tag = "iB1_Seban" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iB1_Hinbn" Then Call Look_Hinbn: Exit Sub
End Sub
'+----------------------------+
'+       取引先検索            +
'+----------------------------+
Private Sub Look_Torcd()

    RV_Left = 0:    RV_Top = 0
    RV_TorKb = 3: RV_Call = "C"
    
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
    RV_TorKb = 3
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
'+       品番検索            +
'+----------------------------+
Private Sub Look_Hinbn()

    RV_Left = 0: RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 0
    RVI_Torcd = RTrim(iH1_Torcd)
    RVI_Ukeir = RTrim(iH1_Ukeir)
    RVI_Seban = RTrim(iB1_Seban)
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
       iB1_Seban = RTrim(RV_Seban)
       iB1_Hinbn = RTrim(RV_Hinbn)
       Call Hinbn_Set
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
        If TorNmGet(iH1_Torcd, 3) Then
            H1lb_Tornm = RTrim(TRM.略称)
        End If
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
'+  iH1_Ymd
'+-----------------------------+
Private Sub iH1_Ymd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Ymd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Nouno
'+-----------------------------+
Private Sub iB1_Nouno_GotFocus()
    Set ActObj = iB1_Nouno
    B_TabDummy.Tag = "S"
End Sub
Private Sub iB1_Nouno_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
'+-----------------------------+
'+  iB1_Seban
'+-----------------------------+
Private Sub iB1_Seban_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Set ActObj = iB1_Seban
'    B_TabDummy.Tag = "S"
End Sub
Private Sub iB1_Seban_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Seban_LostFocus()
    vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 1) = iB1_Seban
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    Call Seban_Set
End Sub
'+-----------------------------+
'+  iB1_Hinbn
'+-----------------------------+
Private Sub iB1_Hinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Set ActObj = iB1_Hinbn
End Sub
Private Sub iB1_Hinbn_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Hinbn_LostFocus()
    vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 2) = iB1_Hinbn
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    Call Hinbn_Set
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
        vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 4) = iB1_Maisu
        If CisFun.Val2(B1lb_Syuyo) <> 0 Then
            gLong = (CisFun.Val2(B1lb_Syuyo) * CisFun.Val2(iB1_Maisu))
            iB1_HSuryo = Format(gLong, "####")
        End If
    Else
        vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 4) = ""
    End If
    B1lb_Hasu = ""
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
        vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = iB1_HSuryo
    Else
        vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = ""
    End If
    Call Suryo_Set
End Sub
'+-----------------------------+
'+  B1Chk_RBKbn
'+-----------------------------+
Private Sub B1Chk_RBKbn_Click()
    If B1Chk_RBKbn.Value = 0 Then
        B1Chk_RBKbn.Caption = "黒"
        B1Chk_RBKbn.ForeColor = &H80000012
    Else
        B1Chk_RBKbn.Caption = "赤"
        B1Chk_RBKbn.ForeColor = &HFF&
    End If
End Sub
'+-----------------------------+
'+  B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
    If ProcHB = "B1" Then
        If B_TabDummy.Tag = "S" Then
            Call TextToVsGrid(Val(Input_Area.Tag))
            Call Grid_RowChange(vbKeyUp)
        End If
        If B_TabDummy.Tag = "E" Then
            Call TextToVsGrid(Val(Input_Area.Tag))
            Call Grid_RowChange(vbKeyDown)
        End If
        SendKeys "{Tab}"
    End If
End Sub

'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   Dim w_Color     As Long
   w_Color = &H80000008 '初期色ｾｯﾄ(黒)
   Select Case SyoriKb
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
          w_Color = &HFF&
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
       If gObj.Name Like "i*" Then gObj.ForeColor = w_Color
       If gObj.Name Like "*Comb*" Then gObj.ForeColor = w_Color
       If gObj.Name Like "*vsGrid*" Then gObj.ForeColor = w_Color
'      * PF-Key 表示設定 *
       If SyoriKb <> "LOC" Then
          If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
             gObj.Name = "PB_DEL" Then
             gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKb)
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
'        Back_UYmd.Visible = False
    End If
    If HeadBody = "B1" Then
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
        If SyoriKb = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call HeadBodyClear("B")
            ProcHB = "H1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKb)
        Else
            ProcHB = "B1"
            Call DispChange(ProcHB)
            vsGrid.Row = 1
            Call VsGridToText(1)
        End If
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKb)
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
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
'納入先
    If Trim(iH1_Torcd) = "" Then
        CisFun.MB_MSG(2) = "    納入先を入力して下さい            "
        CisFun.MBOX
        iH1_Torcd.SetFocus
        Exit Function
    End If
    If Not TorNmGet(iH1_Torcd, 3) Then
        CisFun.MB_MSG(2) = "    納入先　未登録！（取引先マスタ）            "
        CisFun.MBOX
        iH1_Torcd.SetFocus
        Exit Function
    End If
'受入
    If RTrim(iH1_Ukeir) <> "" Then
        TUM.取引先CD = RTrim(iH1_Torcd)
        TUM.受入 = RTrim(iH1_Ukeir)
        TUM.取引先区分 = 0
        If Not TUMRead("", 1) Then
            CisFun.MB_MSG(2) = "    受入　未登録！（取引先マスタ）            "
            CisFun.MBOX
            iH1_Torcd.SetFocus
            Exit Function
        End If
    End If
'日付チェック
    If RTrim(iH1_Ymd) = "" Then
        CisFun.MB_MSG(1) = "     　【  検収日日 必須！ 】     "
        CisFun.MB_MSG(3) = "       検収日 を入力して下さい。      "
        CisFun.MBOX
        iH1_Ymd.SetFocus
        Exit Function
    End If
    If Not iH1_Ymd.cChkResult Then
        CisFun.MB_MSG(1) = "     　【  検収日 異常！ 】     "
        CisFun.MB_MSG(3) = "       検収日 を再入力して下さい。      "
        CisFun.MBOX
        iH1_Ymd.SetFocus
        Exit Function
    End If
    If RTrim(iH1_Ymd) > Format(Now(), "yyyymmdd") Then
        CisFun.MB_MSG(1) = "     検収日が未来の日付です。     "
        CisFun.MB_MSG(3) = "     　　　よろしいですか？      "
        CisFun.MB_Button = OK_CAN
        If Not CisFun.MBOX Then
            iH1_Ymd.SetFocus
            Exit Function
        End If
        CisFun.MB_MSG(1) = "        "
        CisFun.MB_MSG(3) = "        "
        CisFun.MB_Button = Error
    End If
    gInt = Kdou_Check(iH1_Ymd)
    If gInt = 1 Then
    '休日
        CisFun.MB_MSG(1) = "       検収日  は非稼働日です。     "
        CisFun.MB_MSG(3) = "     　　よろしいですか？      "
        CisFun.MB_Button = OK_CAN
        If Not CisFun.MBOX Then
            iH1_Ymd.SetFocus
            Exit Function
        End If
        CisFun.MB_MSG(1) = "        "
        CisFun.MB_MSG(3) = "        "
        CisFun.MB_Button = Error
    End If
    If gInt = 2 Then
    'カレンダ未登録
        CisFun.MB_MSG(1) = "     " & RTrim(iH1_Ymd.cYear) & "年" & RTrim(iH1_Ymd.cMonth) & "月のカレンダ未登録！     "
        CisFun.MB_MSG(3) = "     　　(　カレンダマスタ　)　"
        CisFun.MBOX
        iH1_Ymd.SetFocus
        Exit Function
    End If
' マスタ入力 SELECT文 ｾｯﾄ
    gSL_Select = " SELECT * "
    gSL_Select = gSL_Select & " FROM 検収テーブル "
    gSL_Select = gSL_Select & " WHERE データ区分 = 1 "
    gSL_Select = gSL_Select & "   AND 納入先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND 受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   AND 検収日 = '" & RTrim(iH1_Ymd) & "'"
    gSL_Select = gSL_Select & "   AND (実績年月 = ''"
    gSL_Select = gSL_Select & "    OR  実績年月 >= '" & Trim(wsURYM) & "')"
    If DBInput("INV") Then
        If SyoriKb = "ADD" Then
            CisFun.MB_MSG(2) = "    今回検収データ　登録済！           "
            CisFun.MBOX
            iH1_Torcd.SetFocus
            Exit Function
        End If
    Else
        If SyoriKb <> "ADD" Then
            CisFun.MB_MSG(2) = "    今回検収データ　未登録！           "
            CisFun.MBOX
            iH1_Torcd.SetFocus
            Exit Function
        End If
    End If
       
    Head1Chk = True
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
'<< 行を詰める >>
    With vsGrid
        For gCnt1 = 1 To .Rows - 1
            If RTrim(.TextMatrix(gCnt1, 1)) = "" And _
                RTrim(.TextMatrix(gCnt1, 2)) = "" And _
                RTrim(.TextMatrix(gCnt1, 3)) = "" Then
                For gCnt2 = gCnt1 + 1 To .Rows - 1
                    If RTrim(.TextMatrix(gCnt2, 1)) <> "" Or _
                        RTrim(.TextMatrix(gCnt2, 2)) <> "" Or _
                        RTrim(.TextMatrix(gCnt2, 3)) <> "" Then
                        .TextMatrix(gCnt1, 1) = .TextMatrix(gCnt2, 1)
                        .TextMatrix(gCnt1, 2) = .TextMatrix(gCnt2, 2)
                        .TextMatrix(gCnt1, 3) = .TextMatrix(gCnt2, 3)
                        .TextMatrix(gCnt1, 4) = .TextMatrix(gCnt2, 4)
                        .TextMatrix(gCnt1, 5) = .TextMatrix(gCnt2, 5)
                        .TextMatrix(gCnt1, 6) = .TextMatrix(gCnt2, 6)
                        .TextMatrix(gCnt1, 7) = .TextMatrix(gCnt2, 7)
                        .TextMatrix(gCnt1, 8) = .TextMatrix(gCnt2, 8)
                        .TextMatrix(gCnt1, 9) = .TextMatrix(gCnt2, 9)
                        .TextMatrix(gCnt1, 10) = .TextMatrix(gCnt2, 10)
                        .TextMatrix(gCnt1, 11) = .TextMatrix(gCnt2, 11)
                        .TextMatrix(gCnt1, 12) = .TextMatrix(gCnt2, 12)
                        .TextMatrix(gCnt1, 13) = .TextMatrix(gCnt2, 13)
                        .TextMatrix(gCnt1, 14) = .TextMatrix(gCnt2, 14)
                        .TextMatrix(gCnt1, 15) = .TextMatrix(gCnt2, 15)
                        .TextMatrix(gCnt1, 16) = .TextMatrix(gCnt2, 16)
                        
                        .TextMatrix(gCnt2, 1) = ""
                        .TextMatrix(gCnt2, 2) = ""
                        .TextMatrix(gCnt2, 3) = ""
                        .TextMatrix(gCnt2, 4) = ""
                        .TextMatrix(gCnt2, 5) = ""
                        .TextMatrix(gCnt2, 6) = ""
                        .TextMatrix(gCnt2, 7) = ""
                        .TextMatrix(gCnt2, 8) = ""
                        .TextMatrix(gCnt2, 9) = ""
                        .TextMatrix(gCnt2, 10) = ""
                        .TextMatrix(gCnt2, 11) = ""
                        .TextMatrix(gCnt2, 12) = ""
                        .TextMatrix(gCnt2, 13) = ""
                        .TextMatrix(gCnt2, 14) = ""
                        .TextMatrix(gCnt2, 15) = ""
                        .TextMatrix(gCnt2, 16) = ""
                        Exit For
                    End If
                Next gCnt2
            End If
        Next gCnt1
'
        For gCnt1 = 1 To .Rows - 1
            If RTrim(.TextMatrix(gCnt1, 1)) = "" And _
                RTrim(.TextMatrix(gCnt1, 2)) = "" And _
                RTrim(.TextMatrix(gCnt1, 3)) = "" Then
                Exit For
            End If
        
            If RTrim(.TextMatrix(gCnt1, 1)) = "" And _
                RTrim(.TextMatrix(gCnt1, 3)) <> "" Then
                CisFun.MB_MSG(1) = "    納入番号を入力して下さい      "
                CisFun.MB_MSG(3) = "      【 " & gCnt1 & "行目 】"
                CisFun.MBOX
                Set ErrObj = iB1_Nouno
'                gCnt1 = gCnt2
                GoTo Body1Chk_Err
            End If
        
        '<< 重複チェック >>
            If RTrim(.TextMatrix(gCnt1, 1)) <> "" Then
            For gCnt2 = gCnt1 + 1 To .Rows - 1
'                If RTrim(.TextMatrix(gCnt1, 1)) = RTrim(.TextMatrix(gCnt2, 1)) And
                If RTrim(.TextMatrix(gCnt1, 1)) = RTrim(.TextMatrix(gCnt2, 1)) And _
                   RTrim(.TextMatrix(gCnt1, 3)) = RTrim(.TextMatrix(gCnt2, 3)) And _
                   RTrim(.TextMatrix(gCnt1, 14)) = RTrim(.TextMatrix(gCnt2, 14)) Then
                    CisFun.MB_MSG(1) = "    納入番号・品番・赤黒区分が重複しています。            "
                    CisFun.MB_MSG(3) = "    " & gCnt1 & "行目と " & gCnt2 & "行目            "
                    CisFun.MBOX
                    Set ErrObj = iB1_Seban
                    gCnt1 = gCnt2
                    GoTo Body1Chk_Err
                End If
            Next gCnt2
            End If
        ' 背番号半角チェック
            If RTrim(RTrim(.TextMatrix(gCnt1, 2))) <> "" And Not HalfSizeChk(RTrim(.TextMatrix(gCnt1, 2))) Then
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
        '<< 品番取引先マスタ存在チェック >>
            gSL_Select = "SELECT * FROM 品番取引先マスタ"
            gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(.TextMatrix(gCnt1, 3)) & "'"
            gSL_Select = gSL_Select & "   AND 背番号 = '" & RTrim(.TextMatrix(gCnt1, 2)) & "'"
            gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(iH1_Torcd) & "'"
            gSL_Select = gSL_Select & "   AND 受入 = '" & RTrim(iH1_Ukeir) & "'"
            gSL_Select = gSL_Select & "   AND 品目 = 0"
            If Not HTMRead(gSL_Select, 1) Then
                CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　品番マスタ　未登録！】            "
                CisFun.MBOX
                Set ErrObj = iB1_Seban
                GoTo Body1Chk_Err
            End If
'''            If OpH1_HKbn(1).Tag = "5" And HTM.SK区分 <> 1 Then
'''                CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　かんばん部品！】            "
'''                CisFun.MBOX
'''                Set ErrObj = iB1_Seban
'''                GoTo Body1Chk_Err
'''            End If
        '<< かんばん部品の時、枚数チェック >>
'            If HTM.SK区分 = 0 And CisFun.Val2(.TextMatrix(gCnt1, 6)) = 0 Then
'                CisFun.MB_MSG(1) = "    　　" & gCnt1 & "行目【　枚数　必須！】            "
'                CisFun.MB_MSG(3) = "    かんばん部品の時、枚数を入力して下さい。            "
'                CisFun.MBOX
'                Set ErrObj = iB1_Maisu
'                GoTo Body1Chk_Err
'            End If
        '<< 検収数チェック >>
'            If CisFun.Val2(.TextMatrix(gCnt1, 8)) = 0 Then
'                CisFun.MB_MSG(1) = "    " & gCnt1 & "行目【　検収数　必須！】            "
'                CisFun.MB_MSG(3) = "    検収数を入力して下さい。            "
'                CisFun.MBOX
'                Set ErrObj = iB1_Maisu
'                GoTo Body1Chk_Err
'            End If
            
        Next gCnt1
    End With
'
'
    Body1Chk = True
    Exit Function
Body1Chk_Err:
    vsGrid.Row = gCnt1
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
    Call HeadBodyClear("B1")
    
    If Not KNTRead(gSL_Select, 0, 1) Then
        Call KNTClose
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
        Call GridSet
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
    ' 検収テーブル
        .SQL = "DELETE FROM 検収テーブル "
        .SQL = .SQL & " WHERE 納入先 = '" & RTrim(iH1_Torcd) & "'"
        .SQL = .SQL & "   And 受入 = '" & RTrim(iH1_Ukeir) & "'"
        .SQL = .SQL & "   And 検収日 = '" & RTrim(iH1_Ymd) & "'"
        .SQL = .SQL & "   And データ区分 = 1"
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
    Dim AddFlg      As Boolean
    Dim wsKanri     As Long
    Dim wsKanriGyo    As Byte
    Dim wsMaisu       As Long
    Dim wsSuryo       As Long


    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        If SyoriKb = "ADD" Then
            .MB_MSG(2) = "     追加処理を行います      "
            .MB_Title = "追加処理"
        Else
            .MB_MSG(2) = "     修正処理を行います      "
            .MB_Title = "修正処理"
        End If
        
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Call VsGridToText(Val(Input_Area.Tag))
            Call GridTextEnabled(True)
            Exit Function
        End If
    End With
    On Error GoTo DBPut_Err
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
'   マスタ項目移送
    For gIndex = 1 To vsGrid.Rows - 1
            
        wsMaisu = CisFun.Val2(vsGrid.TextMatrix(gIndex, 6))
        wsSuryo = CisFun.Val2(vsGrid.TextMatrix(gIndex, 8))
        wsKanri = CisFun.Val2(vsGrid.TextMatrix(gIndex, 10))
        wsKanriGyo = CisFun.Val2(vsGrid.TextMatrix(gIndex, 17))
        
        If wsSuryo = 0 Then
            ' * 検収データ削除 *
            If wsKanri <> 0 Then
                With CisDB
                    .SQL = "DELETE FROM 検収テーブル "
                    .SQL = .SQL & " WHERE 納入先 = '" & RTrim(iH1_Torcd) & "'"
                    .SQL = .SQL & "   And 受入 = '" & RTrim(iH1_Ukeir) & "'"
                    .SQL = .SQL & "   And 検収日 = '" & RTrim(iH1_Ymd) & "'"
                    .SQL = .SQL & "   And 検収管理NO = " & wsKanri
                    .SQL = .SQL & "   And 検収管理行NO = " & wsKanriGyo
                    .SQL = .SQL & "   And データ区分 = 1"
                    .DBExec
                    .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
                End With
                iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")
            End If
        
        Else
        
            '   <><><> 検収管理NOの獲得 <><><>
            If wsKanri = 0 Then
                With CisDB
                    .SQL = "番号獲得"
                    .StoadoCount = 3
                    .ParaNo = 0: .ParaIO = Return用
                    .ParaNo = 1: .ParaIO = Input用
                    .ParaNo = 2: .ParaIO = OutPut用
                    .ParaNo = 3: .ParaIO = Return用
                    .ParaNo = 1: .ParaValue = 13

                    If Not .DBStored Then
                        CisFun.MB_Lines = 4
                        CisFun.MB_MSG(2) = "    検収管理NOの獲得に失敗しました            "
                        CisFun.MB_Title = "採番処理"
                        CisFun.MB_Button = Error
                        CisFun.MBOX
                        Exit Function
                    Else
                        .ParaNo = 2: wsKanri = .ParaValue
                    End If
                    wsKanriGyo = 1
                End With
            End If
            
            Call ItemsClearKNT
            AddFlg = True
            gSL_Select = "SELECT * FROM 検収テーブル "
            gSL_Select = gSL_Select & " WHERE 検収管理NO = " & wsKanri
            gSL_Select = gSL_Select & "   AND 検収管理行NO = " & wsKanriGyo
            If KNTRead(gSL_Select, 1) Then
                AddFlg = False
            End If
            
            If KNT.枚数 <> wsMaisu Or KNT.検収数 <> wsSuryo Then
                With KNT
                    .検収管理NO = wsKanri
                    .検収管理行NO = wsKanriGyo
                    .検収日 = RTrim(iH1_Ymd)
                    .納入先 = RTrim(iH1_Torcd)
                    .受入 = RTrim(iH1_Ukeir)
                    
                    .納入番号 = RTrim(vsGrid.TextMatrix(gIndex, 1))
                    .品番 = RTrim(vsGrid.TextMatrix(gIndex, 3))
                    .背番号 = RTrim(vsGrid.TextMatrix(gIndex, 2))
                    .赤黒区分 = RTrim(vsGrid.TextMatrix(gIndex, 14))
                    .枚数 = wsMaisu
                    .収容数 = CisFun.Val2(vsGrid.TextMatrix(gIndex, 5))
                    .検収数 = wsSuryo
                    .作成種別 = "G"
                    .データ区分 = 1
                End With
    
                If AddFlg Then
                 '( 追加処理 )
'                    KNT.作成者 = RTrim(uINTanto)
                    Call KNTInsert
                    iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
                Else
                ' ( 更新処理 )
'                    KNT.更新者 = RTrim(uINTanto)
                    Call KNTUpdate
                    iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
                End If
            End If
        
        End If
    
    Next gIndex
'
' ( 追加処理 )
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    B_TabDummy.Tag = ""
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'*------------------------------------------*'
'*    背番号による品番・収容数の表示          *
'*------------------------------------------*'
Private Sub Seban_Set()
    If RTrim(iB1_Seban) = "" Then Exit Sub
    Call ItemsClearHTM
    gSL_Select = "SELECT * FROM 品番取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 背番号 = '" & RTrim(iB1_Seban) & "'"
    gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND 受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   AND 品目 = 0"
    Call HTMRead(gSL_Select, 0)
    Call HTMClose
    If CisDB.RecordCount = 0 Then Exit Sub
    If CisDB.RecordCount > 1 Then Exit Sub
    iB1_Hinbn = RTrim(HTM.品番)
'収容数表示
    Call Syuyo_Set
End Sub
'*------------------------------------------*'
'*    背番号による品番・収容数の表示          *
'*------------------------------------------*'
Private Sub Hinbn_Set()
    Call ItemsClearHTM
    gSL_Select = "SELECT * FROM 品番取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iB1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND 受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   AND 品目 = 0"
    If Not HTMRead(gSL_Select, 1) Then Exit Sub
    iB1_Seban = RTrim(HTM.背番号)
    iB1_Hinbn = RTrim(HTM.品番)
'収容数表示
    Call Syuyo_Set
End Sub
'*------------------------------------------*'
'*    収容数の表示                           *
'*------------------------------------------*'
Private Sub Syuyo_Set()
    Dim wNYmd   As String
    Dim wNBin   As String
    
    wNYmd = RTrim(iH1_Ymd)
'    mNYmd = RTrim(iH1_Ymd)
'    mNBin = RTrim(iH1_Bin)
    If RTrim(HTM.変更日) <> "" Then
        If RTrim(HTM.変更日) > RTrim(wNYmd) Then
            B1lb_Syuyo = Format(HTM.収容数1, "#,###")
        Else
            B1lb_Syuyo = Format(HTM.収容数2, "#,###")
        End If
    Else
        B1lb_Syuyo = Format(HTM.収容数1, "#,###")
    End If
    Call Suryo_Set
End Sub
'*------------------------------------------*'
'*    枚数・端数・発注数の再表示              *
'*------------------------------------------*'
Private Sub Suryo_Set()
    B1lb_Hasu = ""
    If CisFun.Val2(B1lb_Syuyo) <> 0 Then
        If RTrim(iB1_HSuryo) = "" Then
            If CisFun.Val2(iB1_Maisu) <> 0 Then
                gLong = (CisFun.Val2(B1lb_Syuyo) * CisFun.Val2(iB1_Maisu))
                iB1_HSuryo = Format(gLong, "####")
            End If
        Else
            gCnt1 = CisFun.Val2(iB1_HSuryo) \ CisFun.Val2(B1lb_Syuyo)
            gCnt2 = CisFun.Val2(iB1_HSuryo) Mod CisFun.Val2(B1lb_Syuyo)
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
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 左中, あり, 10, 1, "納入番号")
        '*-*-*-*-* 2008/04/10 Update Start
        'Call .FixedSet(中中, 左中, あり, 7, 2, "背番号")
        Call .FixedSet(中中, 左中, あり, 9, 2, "背番号")
        '*-*-*-*-* 2008/04/10 Update End
        Call .FixedSet(中中, 左中, あり, 25, 3, "品　番")
        Call .FixedSet(中中, 中中, あり, 5, 4, "赤/黒")
        Call .FixedSet(中中, 右中, あり, 8, 5, "収容数")
        Call .FixedSet(中中, 右中, あり, 7, 6, "枚　数")
        Call .FixedSet(中中, 右中, あり, 7, 7, "端　数")
        Call .FixedSet(中中, 右中, あり, 9, 8, "検収数")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 9, "X")
        Call .FixedSet(中中, 左中, なし, 0, 10, "検収管理NO")
        Call .FixedSet(中中, 左中, なし, 0, 11, "検収日")
        Call .FixedSet(中中, 左中, なし, 0, 12, "納入先")
        Call .FixedSet(中中, 左中, なし, 0, 13, "受入")
        Call .FixedSet(中中, 左中, なし, 0, 14, "赤黒区分")
        Call .FixedSet(中中, 左中, なし, 0, 15, "実績年月")
        Call .FixedSet(中中, 左中, なし, 0, 16, "変更前数量")
        Call .FixedSet(中中, 左中, なし, 0, 17, "検収管理行NO")
        Call .FixedSet(中中, 左中, なし, 0, 18, "Select")
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
    Do Until Not KNT_RDSTS
        'Grid Row,Col
        ID = ID + 1
        With vsGrid
            .TextMatrix(ID, GridCol(0)) = ID
            .TextMatrix(ID, GridCol(1)) = RTrim(KNT.納入番号)
            .TextMatrix(ID, GridCol(2)) = RTrim(KNT.背番号)
            .TextMatrix(ID, GridCol(3)) = RTrim(KNT.品番)
            
            If KNT.赤黒区分 = 0 Then
                .TextMatrix(ID, GridCol(4)) = "黒"
            Else
                .TextMatrix(ID, GridCol(4)) = "赤"
            End If
            .TextMatrix(ID, GridCol(14)) = RTrim(KNT.赤黒区分)
            .TextMatrix(ID, GridCol(5)) = CisFun.RSetFld(KNT.収容数, 7, "#,###")
            .TextMatrix(ID, GridCol(6)) = CisFun.RSetFld(KNT.枚数, 7, "#,###")

'< 2005/12/14 MNT START >
'
'            gLong = KNT.検収数 Mod KNT.収容数
            gLong = 0
            If KNT.収容数 <> 0 Then gLong = KNT.検収数 Mod KNT.収容数
'
'< 2005/12/14 MNT END >

            .TextMatrix(ID, GridCol(7)) = CisFun.RSetFld(gLong, 7, "#,###")
            .TextMatrix(ID, GridCol(8)) = CisFun.RSetFld(KNT.検収数, 9, "#,###")
            .TextMatrix(ID, GridCol(16)) = CisFun.RSetFld(KNT.検収数, 9, "#,###")
            
            If KNT.実績年月 <> "" And KNT.実績年月 < wsURYM Then
                .TextMatrix(ID, GridCol(9)) = "X"
            End If
            .TextMatrix(ID, GridCol(10)) = KNT.検収管理NO
            .TextMatrix(ID, GridCol(11)) = KNT.検収日
            .TextMatrix(ID, GridCol(12)) = KNT.納入先
            .TextMatrix(ID, GridCol(13)) = KNT.受入
            .TextMatrix(ID, GridCol(15)) = "X"
            .TextMatrix(ID, GridCol(17)) = KNT.検収管理行NO
        End With
        Call KNTReadNext(1)
    Loop

    With vsGrid
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    Call KNTClose(1)
    
End Sub
'==========================================================='
'   入力テキスト表示・非表示
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
    Input_Area.Visible = EnabledFlg
End Sub
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub

'テキストからグリッドに退避
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
            
'--------------------------------------------------------------< 2006/06/06 MNT >
'            If RTrim(vsGrid.TextMatrix(gInt, 2)) <> "" Then
            If RTrim(vsGrid.TextMatrix(gInt, 3)) <> "" Then
'--------------------------------------------------------------< 2006/06/06 MNT END >
                
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
'グリッドからテキストに表示
    Call VsGridToText(vsGrid.Row)
    If Not B1_Back.Enabled Then
        If ActObj.Name = "iB1_Nouno" Or _
           ActObj.Name = "iB1_Seban" Or _
           ActObj.Name = "iB1_Hinbn" Then
            iB1_Maisu.SetFocus
        End If
    End If
    
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
        .TextMatrix(tRow, 1) = RTrim(iB1_Nouno)
        .TextMatrix(tRow, 2) = RTrim(iB1_Seban)
        .TextMatrix(tRow, 3) = RTrim(iB1_Hinbn)
        
        If RTrim(iB1_Hinbn) <> "" Then
            .TextMatrix(tRow, 14) = B1Chk_RBKbn.Value
            If B1Chk_RBKbn.Value = 0 Then
                .TextMatrix(tRow, 4) = "黒"
            Else
                .TextMatrix(tRow, 4) = "赤"
            End If
        End If
        .TextMatrix(tRow, 5) = Format(CisFun.Val2(B1lb_Syuyo), "#,###")
        
        Call Suryo_Set
        If CisFun.Val2(iB1_Maisu) <> 0 Then
            .TextMatrix(tRow, 6) = Format(CisFun.Val2(iB1_Maisu), "#,###")
        Else
            .TextMatrix(tRow, 6) = ""
        End If
        If CisFun.Val2(iB1_HSuryo) <> 0 Then
            .TextMatrix(tRow, 8) = Format(CisFun.Val2(iB1_HSuryo), "#,###")
        Else
            .TextMatrix(tRow, 8) = ""
        End If
        
        .TextMatrix(tRow, 7) = Format(CisFun.Val2(B1lb_Hasu), "#,###")
        
        
    End With
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    Static wsToText     As Boolean
    
    If wsToText Then Exit Sub
    wsToText = True
    
    With vsGrid
        Input_Area.Enabled = True
        Input_Area.Tag = tRow
        Input_Area.Move Input_Area.Left, vsGrid.CellTop, Input_Area.Width, vsGrid.CellHeight
        
        '納入番号
        iB1_Nouno = RTrim(.TextMatrix(tRow, 1))
        '背番号
        iB1_Seban = RTrim(.TextMatrix(tRow, 2))
        '品番
        iB1_Hinbn = RTrim(.TextMatrix(tRow, 3))
        '赤黒
        B1Chk_RBKbn.Value = CisFun.Val2(.TextMatrix(tRow, 14))
        '収容数
        B1lb_Syuyo = Format(CisFun.Val2(.TextMatrix(tRow, 5)), "#,###")
        '枚数
        iB1_Maisu = CisFun.Val2(.TextMatrix(tRow, 6))
        '端数
        B1lb_Hasu = Format(CisFun.Val2(.TextMatrix(tRow, 7)), "#,###")
        '検収数
        iB1_HSuryo = CisFun.Val2(.TextMatrix(tRow, 8))
    
        If RTrim(vsGrid.TextMatrix(tRow, 9)) <> "" Then
            Input_Area.Enabled = False
        End If
        If CisFun.Val2(vsGrid.TextMatrix(tRow, 10)) <> 0 Then
            iB1_Seban.Enabled = False
            iB1_Hinbn.Enabled = False
        Else
            iB1_Seban.Enabled = True
            iB1_Hinbn.Enabled = True
        End If
    End With
    
    
    wsToText = False
End Sub
Private Sub VSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
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
'Private Sub VSGrid_DblClick()
'    If RTrim(vsGrid.TextMatrix(vsGrid.Row, 1)) = "" And RTrim(vsGrid.TextMatrix(vsGrid.Row, 2)) = "" Then
'        Exit Sub
'    End If
'    With CisFun
'        .MB_Lines = 4
'        .MB_MSG(2) = "     " & vsGrid.Row & " 行目をクリアしますか？      "
'        .MB_Title = ""
'        .MB_Button = OK_CAN
'        If Not .MBOX Then Exit Sub
'    End With
'    iB1_Seban = "": iB1_Hinbn = "": B1lb_Syuyo = ""
'    iB1_Maisu = "": iB1_HSuryo = "": B1lb_Hasu = ""
'    B1Chk_RBKbn.Value = 0
'    Call TextToVsGrid(Val(Input_Area.Tag))
'
'End Sub
Private Sub vsGrid_Scroll()
    'スクロール時はテキストボックスを非表示にします。
'    Call GridTextEnabled(False)
End Sub
