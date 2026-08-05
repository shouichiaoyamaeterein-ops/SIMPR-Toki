VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CTP3310 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "新e-かんばん 前工程情報出力"
   ClientHeight    =   9705
   ClientLeft      =   480
   ClientTop       =   270
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
   ScaleHeight     =   9705
   ScaleWidth      =   14790
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D lb_TxtOut 
      Height          =   390
      Left            =   1245
      Top             =   2295
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   688
      BackColor       =   16777152
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4XXXXXXXXX5"
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
      cFont3D         =   2
      cAlingnment     =   1
      cBoderStyle     =   1
      cPositionX      =   60
   End
   Begin Cis3D_v60.CIS3D CIS3D8 
      Height          =   390
      Left            =   165
      Top             =   2295
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   688
      Caption         =   "出力先"
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
   Begin VB.PictureBox H1_Area1 
      Height          =   1035
      Left            =   420
      ScaleHeight     =   975
      ScaleWidth      =   7890
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   705
      Width           =   7950
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   465
         Left            =   5355
         Top             =   495
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   820
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
      Begin Cis3D_v60.CIS3D CIS3D16 
         Height          =   465
         Left            =   2325
         Top             =   30
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   820
         ForeColor       =   16711680
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
         Begin CisText_V60.CisText iH1_NTorcd 
            Height          =   375
            Left            =   825
            TabIndex        =   2
            Top             =   45
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            cWildeCard      =   -1  'True
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
         Begin Cis3D_v60.CIS3D H1lb_NTornm 
            Height          =   360
            Left            =   1710
            Top             =   45
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
         Begin CisText_V60.CisText iH1_NTorcdKJ 
            Height          =   375
            Left            =   1440
            TabIndex        =   3
            Top             =   45
            Width           =   255
            _ExtentX        =   450
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
            Text            =   "X"
            MaxLength       =   1
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   465
         Left            =   15
         Top             =   30
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   820
         ForeColor       =   16711680
         Caption         =   "納入先事業体"
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
         Begin CisText_V60.CisText iH1_NJigyo 
            Height          =   375
            Left            =   1560
            TabIndex        =   0
            Top             =   45
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   661
            cWildeCard      =   -1  'True
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
         Begin CisText_V60.CisText iH1_NJigyoKJ 
            Height          =   375
            Left            =   1995
            TabIndex        =   1
            Top             =   45
            Width           =   255
            _ExtentX        =   450
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
            Text            =   "X"
            MaxLength       =   1
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D15 
         Height          =   465
         Index           =   12
         Left            =   6570
         Top             =   30
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   820
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin CisText_V60.CisText iH1_Uke 
            Height          =   375
            Left            =   795
            TabIndex        =   4
            Top             =   45
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   661
            cWildeCard      =   -1  'True
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
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   465
         Left            =   15
         Top             =   495
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   820
         ForeColor       =   16711680
         Caption         =   "仕入先事業体"
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
         Begin CisText_V60.CisText iH1_SJigyo 
            Height          =   375
            Left            =   1560
            TabIndex        =   5
            Top             =   45
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   661
            cWildeCard      =   -1  'True
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
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   465
         Left            =   2325
         Top             =   495
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   820
         ForeColor       =   16711680
         Caption         =   "仕入先"
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
         Begin CisText_V60.CisText iH1_STorcd 
            Height          =   375
            Left            =   825
            TabIndex        =   6
            Top             =   45
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            cWildeCard      =   -1  'True
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
         Begin CisText_V60.CisText iH1_STorcdKJ 
            Height          =   375
            Left            =   1440
            TabIndex        =   7
            Top             =   45
            Width           =   255
            _ExtentX        =   450
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
            Text            =   "X"
            MaxLength       =   1
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D15 
         Height          =   465
         Index           =   13
         Left            =   4065
         Top             =   495
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   820
         ForeColor       =   16711680
         Caption         =   "出荷場"
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
         Begin CisText_V60.CisText iH1_Syukaba 
            Height          =   375
            Left            =   780
            TabIndex        =   8
            Top             =   45
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   661
            cWildeCard      =   -1  'True
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
      Left            =   11145
      ScaleHeight     =   735
      ScaleWidth      =   3510
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1410
      Width           =   3570
      Begin CisBtn_60.CisBtn PB_CAN 
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
      TabIndex        =   10
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
      Height          =   6960
      Left            =   165
      ScaleHeight     =   6900
      ScaleWidth      =   14490
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2700
      Width           =   14550
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   6765
         Index           =   0
         Left            =   60
         Top             =   60
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   11933
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
            Height          =   6690
            Index           =   1
            Left            =   30
            Top             =   30
            Width           =   14340
            _ExtentX        =   25294
            _ExtentY        =   11800
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
               Height          =   6555
               Left            =   45
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   60
               Width           =   14220
               _Version        =   196608
               _ExtentX        =   25082
               _ExtentY        =   11562
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
               FormatString    =   "  |事業体|納入先ｺｰﾄﾞ|受入|納入先名称           |自工程  |事業体|契約先名称           |仕入先ｺｰﾄﾞ|出荷場|前工程  |件数||||||||"
               Rows            =   18
               Cols            =   20
               FixedRows       =   2
               BackColor       =   12648447
               BackColorFixed  =   12632256
               BackColorBkg    =   12632256
               FocusRect       =   0
               HighLight       =   2
               RowHeightMin    =   360
               GridLines       =   2
               SelectionMode   =   1
               Begin Cis3D_v60.CIS3D CIS3D2 
                  Height          =   45
                  Left            =   15
                  Top             =   330
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   79
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
                  cBoderColor1    =   12632256
                  cBoderColor2    =   12632256
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cBoderStyle     =   2
               End
               Begin Cis3D_v60.CIS3D CIS3D7 
                  Height          =   720
                  Left            =   13230
                  Top             =   0
                  Width           =   570
                  _ExtentX        =   1005
                  _ExtentY        =   1270
                  BackColor       =   16777152
                  ForeColor       =   0
                  Caption         =   "対象件数"
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
                  cPositionX      =   30
                  cPositionY      =   129
               End
               Begin Cis3D_v60.CIS3D CIS3D3 
                  Height          =   360
                  Index           =   0
                  Left            =   330
                  Top             =   360
                  Width           =   810
                  _ExtentX        =   1429
                  _ExtentY        =   635
                  BackColor       =   16761024
                  Caption         =   "事業体"
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
                  cPositionX      =   20
               End
               Begin Cis3D_v60.CIS3D CIS3D1 
                  Height          =   360
                  Index           =   0
                  Left            =   330
                  Top             =   0
                  Width           =   6330
                  _ExtentX        =   11165
                  _ExtentY        =   635
                  BackColor       =   16761024
                  Caption         =   "納   入   先"
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
               Begin Cis3D_v60.CIS3D CIS3D3 
                  Height          =   360
                  Index           =   1
                  Left            =   1140
                  Top             =   360
                  Width           =   1290
                  _ExtentX        =   2275
                  _ExtentY        =   635
                  BackColor       =   16761024
                  Caption         =   "納入先ｺｰﾄﾞ"
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
                  cPositionX      =   20
               End
               Begin Cis3D_v60.CIS3D CIS3D3 
                  Height          =   360
                  Index           =   2
                  Left            =   2430
                  Top             =   360
                  Width           =   570
                  _ExtentX        =   1005
                  _ExtentY        =   635
                  BackColor       =   16761024
                  Caption         =   "受入"
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
                  cPositionX      =   20
               End
               Begin Cis3D_v60.CIS3D CIS3D3 
                  Height          =   360
                  Index           =   3
                  Left            =   3000
                  Top             =   360
                  Width           =   2610
                  _ExtentX        =   4604
                  _ExtentY        =   635
                  BackColor       =   16761024
                  Caption         =   "納入先名称"
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
                  cPositionX      =   20
               End
               Begin Cis3D_v60.CIS3D CIS3D1 
                  Height          =   360
                  Index           =   1
                  Left            =   6660
                  Top             =   0
                  Width           =   6570
                  _ExtentX        =   11589
                  _ExtentY        =   635
                  BackColor       =   16761087
                  Caption         =   "仕   入   先"
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
               Begin Cis3D_v60.CIS3D CIS3D3 
                  Height          =   360
                  Index           =   4
                  Left            =   6660
                  Top             =   360
                  Width           =   810
                  _ExtentX        =   1429
                  _ExtentY        =   635
                  BackColor       =   16761087
                  Caption         =   "事業体"
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
                  cPositionX      =   20
               End
               Begin Cis3D_v60.CIS3D CIS3D3 
                  Height          =   360
                  Index           =   5
                  Left            =   10095
                  Top             =   360
                  Width           =   1290
                  _ExtentX        =   2275
                  _ExtentY        =   635
                  BackColor       =   16761087
                  Caption         =   "仕入先ｺｰﾄﾞ"
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
                  cPositionX      =   20
               End
               Begin Cis3D_v60.CIS3D CIS3D3 
                  Height          =   360
                  Index           =   6
                  Left            =   11385
                  Top             =   360
                  Width           =   795
                  _ExtentX        =   1402
                  _ExtentY        =   635
                  BackColor       =   16761087
                  Caption         =   "出荷場"
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
                  cPositionX      =   20
               End
               Begin Cis3D_v60.CIS3D CIS3D3 
                  Height          =   360
                  Index           =   7
                  Left            =   7470
                  Top             =   360
                  Width           =   2625
                  _ExtentX        =   4630
                  _ExtentY        =   635
                  BackColor       =   16761087
                  Caption         =   "事業体名称(契約先)"
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
                  cPositionX      =   20
               End
               Begin Cis3D_v60.CIS3D CIS3D3 
                  Height          =   360
                  Index           =   8
                  Left            =   5610
                  Top             =   360
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   635
                  BackColor       =   16761024
                  Caption         =   "自工程"
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
                  cPositionX      =   20
               End
               Begin Cis3D_v60.CIS3D CIS3D3 
                  Height          =   360
                  Index           =   9
                  Left            =   12180
                  Top             =   360
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   635
                  BackColor       =   16761087
                  Caption         =   "前工程"
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
                  cPositionX      =   20
               End
               Begin Cis3D_v60.CIS3D Proc_MSG 
                  Height          =   1260
                  Left            =   3390
                  Top             =   1950
                  Width           =   7185
                  _ExtentX        =   12674
                  _ExtentY        =   2223
                  ForeColor       =   128
                  Caption         =   "前工程情報データ作成中"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   24
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cBoderWidth     =   3
                  cAlingnment     =   7
                  cPositionY      =   -60
                  Begin VB.PictureBox Bar 
                     Height          =   360
                     Left            =   45
                     ScaleHeight     =   300
                     ScaleWidth      =   7035
                     TabIndex        =   14
                     Top             =   855
                     Width           =   7095
                  End
               End
            End
         End
      End
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   -15
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 新e-かんばん 前工程情報出力  】"
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
   Begin Cis3D_v60.CIS3D CIS3D10 
      Height          =   345
      Left            =   10515
      Top             =   2340
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   609
      BackColor       =   12632319
      Caption         =   "表示件数"
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
      cPositionX      =   90
   End
   Begin Cis3D_v60.CIS3D lb_DispSu 
      Height          =   345
      Left            =   11655
      Top             =   2340
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   609
      BackColor       =   12632319
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cAlingnment     =   4
      cBoderStyle     =   1
      cPositionX      =   -60
   End
   Begin Cis3D_v60.CIS3D CIS3D6 
      Height          =   345
      Left            =   12675
      Top             =   2340
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   609
      BackColor       =   12640511
      Caption         =   "選択件数"
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
      cPositionX      =   90
   End
   Begin Cis3D_v60.CIS3D lb_ChoiceSu 
      Height          =   345
      Left            =   13815
      Top             =   2340
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   609
      BackColor       =   12640511
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cAlingnment     =   4
      cBoderStyle     =   1
      cPositionX      =   -60
   End
End
Attribute VB_Name = "CTP3310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   新e-かんばん 前工程情報ﾃｷｽﾄ出力
'**       フォームID    :   CTP3310
'**       処理概要      :
'**
'**       作  成  日    :   2008/07/17  By CIS SIMPR-A
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
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
    
    Dim OutFNo                  As Byte
    Dim CrtTxt                  As String
    Dim CrtVal                  As String
    Dim CrtCnt                  As Long
    
    Dim wCnt                    As Byte
    Dim wMax                    As Byte
    Dim wOki(3)                 As String       ' 所番地
    Dim wSyuyo(3)               As String       ' 収容数
    Dim wHako(3)                As String       ' 箱種
    Dim wDate(3)                As String       ' 変更日
    Dim wBin(3)                 As String       ' 変更便

    Dim CisSimplePrint As New CisSimplePrint

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
        .SelectCol = 18
        Set .SelectDispObj = lb_ChoiceSu
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

    Proc_MSG.Visible = False

'   #-----------------------------#
'   #    前工程情報初期設定取得   #
'   #-----------------------------#
    Call eKanban_Maekotei_IniGet
    lb_TxtOut = RTrim(eKanban_MaeKoteiFld) & RTrim(eKanban_MaeKoteiFile)
    
    
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    
    If RTrim(eKanban_SBasyoCD) = "" Then
        With CisFun
                .MB_Lines = 4
                .MB_MSG(1) = "            設置場所コード　未登録！！                "
                .MB_MSG(3) = "    【新eかんばん連携 環境設定】を行って下さい        "
                .MB_Title = "環境設定"
                .MB_Button = Error
                .MBOX
        End With
        Call PB_END_Click
    End If
    
    If RTrim(eKanban_MachineNo) = "" Then
        With CisFun
                .MB_Lines = 4
                .MB_MSG(1) = "               マシン№　未登録！！                   "
                .MB_MSG(3) = "    【新eかんばん連携 環境設定】を行って下さい        "
                .MB_Title = "環境設定"
                .MB_Button = Error
                .MBOX
        End With
        Call PB_END_Click
    End If
    
    
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
    Call CisFrm.MousePT(11)
    If ProcHB = "B1" Then
       ProcHB = "H1"
       Call DispChange(ProcHB)
       Call HeadBodyClear("B")
       GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_NTorcd" Then Call Look_Torcd(PB_Look.Tag): Exit Sub
    If PB_Look.Tag = "iH1_NTorcdKJ" Then Call Look_Torcd(PB_Look.Tag): Exit Sub
    If PB_Look.Tag = "iH1_Uke" Then Call Look_Ukeir: Exit Sub
End Sub
'+----------------------------+
'+       取引先検索            +
'+----------------------------+
Private Sub Look_Torcd(TextName As String)
    RV_Left = 0
    RV_Top = 0

    RV_Call = "C"
    RV_TorKb = 3
    '
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
       iH1_NTorcd = RV_Torcd
       iH1_NTorcdKJ = RV_Kojyo
       H1lb_NTornm = RV_TorRName
       If TextName = "iH1_NTorcd" Then
          iH1_NTorcd.SetFocus
       Else
          iH1_NTorcdKJ.SetFocus
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
    RV_TorKb = 3            '【受注】納入先/契約先検索
    RV_Torcd = iH1_NTorcd
    RV_Kojyo = ""

    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_NTorcd = RV_Torcd
        iH1_NTorcdKJ = RV_Kojyo
        H1lb_NTornm = RV_TorRName
        iH1_Uke = RV_Ukeir
    End If
End Sub
'+-----------------------------+
'+      iH1_NTorcd
'+-----------------------------+
Private Sub iH1_NTorcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_NTorcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H1lb_NTornm = ""
    If RTrim(iH1_NTorcd) <> "" Then
        If Get_Torcd(RTrim(iH1_NTorcd) & RTrim(iH1_NTorcdKJ), 0) Then
            H1lb_NTornm = TRM_RName
        End If
    End If
End Sub
'+-----------------------------+
'+      iH1_NTorcdKJ
'+-----------------------------+
Private Sub iH1_NTorcdKJ_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_NTorcdKJ_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H1lb_NTornm = ""
    If RTrim(iH1_NTorcd) <> "" Then
        If Get_Torcd(RTrim(iH1_NTorcd) & RTrim(iH1_NTorcdKJ), 0) Then
            H1lb_NTornm = TRM_RName
        End If
    End If
End Sub
'+-----------------------------+
'+      iH1_Uke
'+-----------------------------+
Private Sub iH1_Uke_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Uke_LostFocus()
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
    Dim H2Mode      As Boolean
    Dim H2Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long

    H1Mode = False: H1Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1":   H1Mode = True:  H1Color = gPGotSel
       Case "B1":   B1Mode = True:  B1Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode:  H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode:  B1_Area1.BackColor = B1Color

    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode

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
            If .Name Like "i" & wbc_Enm & "_*" Then .Text = ""
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False: .BackColor = mColor_Off
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Cmb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
        CisVsGrid.Clear
        lb_DispSu = "": lb_ChoiceSu = ""
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
        If Not TXTPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        Call HeadBodyClear("H1")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        
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

' マスタ入力 SELECT文 ｾｯﾄ
    gSL_Select = "SELECT"
    gSL_Select = gSL_Select & " TR1.納入先事業体 , TR1.納入先事業体工区 , TR1.取引先 納入先コード, TR1.工場CD 納入先工区 , " & vbCrLf
    gSL_Select = gSL_Select & "HT.受入 , TR1.略称 納入先略称 , " & vbCrLf
    gSL_Select = gSL_Select & "TR2.仕入先事業体 , TR2.略称 契約先略称 , " & vbCrLf
    gSL_Select = gSL_Select & "SUBSTRING(TR2.被自社,1,4) 仕入先コード , SUBSTRING(TR2.被自社工場,1,1) 仕入先工区 , " & vbCrLf
    gSL_Select = gSL_Select & "ISNULL(ME.自工程,'') 自工程 , ISNULL(ME.前工程,'') 前工程 , " & vbCrLf
    gSL_Select = gSL_Select & "ISNULL(HT.出荷場 ,'') 出荷場 , COUNT(HT.品番) 対象件数 " & vbCrLf
    gSL_Select = gSL_Select & "FROM 品番取引先マスタ HT " & vbCrLf
    gSL_Select = gSL_Select & "LEFT OUTER JOIN 取引先マスタ TR1 " & vbCrLf
    gSL_Select = gSL_Select & "        ON  HT.取引先 = TR1.取引先CD " & vbCrLf
    gSL_Select = gSL_Select & "        AND TR1.取引先区分 = 0 " & vbCrLf
    gSL_Select = gSL_Select & "LEFT OUTER JOIN 取引先マスタ TR2 " & vbCrLf
    gSL_Select = gSL_Select & "        ON  HT.契約先 = TR2.取引先CD " & vbCrLf
    gSL_Select = gSL_Select & "        AND TR2.取引先区分 = 0 " & vbCrLf
    gSL_Select = gSL_Select & "LEFT OUTER JOIN 新TOPPS前工程情報テーブル ME " & vbCrLf
    gSL_Select = gSL_Select & "        ON  ME.納入先 = HT.取引先 " & vbCrLf
    gSL_Select = gSL_Select & "        AND ME.受入 = HT.受入 " & vbCrLf
    gSL_Select = gSL_Select & "        AND ME.契約先 = HT.契約先 " & vbCrLf
    
    With CisFun
         .Where_CreateKB = NewDocument
            
         ' 納入先事業体
         .Where_Create AlphaMe, "ISNULL(TR1.納入先事業体,'')", iH1_NJigyo
         ' 納入先事業体工区
         .Where_Create AlphaMe, "ISNULL(TR1.納入先事業体工区,'')", iH1_NJigyoKJ
         ' 納入先
         .Where_Create AlphaMe, "TR1.取引先", iH1_NTorcd
         ' 納入先工区
         .Where_Create AlphaMe, "TR1.工場CD", iH1_NTorcdKJ
         ' 受入
         .Where_Create AlphaMe, "HT.受入", iH1_Uke
         ' 仕入先事業体
         .Where_Create AlphaMe, "ISNULL(TR2.仕入先事業体,'')", iH1_SJigyo
         ' 仕入先(契約先：被自社)
         .Where_Create AlphaMe, "ISNULL(SUBSTRING(TR2.被自社,1,4),'')", iH1_STorcd
         ' 仕入先工区(契約先：被自社工場)
         .Where_Create AlphaMe, "ISNULL(SUBSTRING(TR2.被自社工場,1,1),'')", iH1_STorcdKJ
         ' 出荷場
         .Where_Create AlphaMe, "ISNULL(HT.出荷場,'')", iH1_Syukaba
    
         If RTrim(.Where_Phrase) <> "" Then
            gSL_Select = gSL_Select & .Where_Phrase & " AND HT.品目 = 0 "
         Else
            gSL_Select = gSL_Select & " WHERE HT.品目 = 0 "
         End If
    End With
    
    'gSL_Select = gSL_Select & "AND HT.品目 = 0 " & vbCrLf
    gSL_Select = gSL_Select & "AND ISNULL(TR1.取引先CD,'') <> '' " & vbCrLf
    gSL_Select = gSL_Select & "AND ISNULL(TR2.取引先CD,'') <> '' " & vbCrLf
    gSL_Select = gSL_Select & "AND ISNULL(TR1.納入先事業体,'') <> '' " & vbCrLf
    gSL_Select = gSL_Select & "AND ISNULL(TR1.納入先事業体工区,'') <> '' " & vbCrLf
    gSL_Select = gSL_Select & "AND ISNULL(TR2.仕入先事業体,'') <> '' " & vbCrLf
    gSL_Select = gSL_Select & "AND ISNULL(TR2.被自社,'') <> '' " & vbCrLf
    gSL_Select = gSL_Select & "GROUP BY TR1.納入先事業体 , TR1.納入先事業体工区 , TR1.取引先 , TR1.工場CD , " & vbCrLf
    gSL_Select = gSL_Select & "         HT.受入 , TR1.略称 , TR2.仕入先事業体 , TR2.略称 , " & vbCrLf
    gSL_Select = gSL_Select & "         SUBSTRING(TR2.被自社,1,4) , SUBSTRING(TR2.被自社工場,1,1) , " & vbCrLf
    gSL_Select = gSL_Select & "         ISNULL(ME.自工程,'') , ISNULL(ME.前工程,'') , ISNULL(HT.出荷場 ,'') " & vbCrLf
    
    If Not DBInput("INV") Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    該当データが存在しません｡            "
            .MB_MSG(3) = "    再度抽出条件を入力してください       "
            .MB_Title = "抽出処理"
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If

Head1Chk_Ed:
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False

    If CisFun.Val2(lb_ChoiceSu) = 0 Then
        With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    出力データを選択して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
        End With
        Exit Function
    End If

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

    If GridSet Then
        DBInput = True
    End If

DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'**************************************
'*                                    *
'*          テキスト出力              *
'*                                    *
'**************************************
'+----------------------+
'+    データ出力処理    +
'+----------------------+
Private Function TXTPut()
    TXTPut = False

' ( 確認メッセージ )
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "    テキストファイル出力を行います。            "
        .MB_Title = "出力確認"
        .MB_MSG(4) = "     よろしいですか？            "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Exit Function
        End If
    End With

    On Error GoTo TXTPut_Err
    
' ( フォルダ確認 )
    
    If Dir(eKanban_MaeKoteiFld, vbDirectory) = "" Then
        DoEvents
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　 出力フォルダが存在しません。        "
            .MB_MSG(4) = "   　 フォルダを作成しますか？            "
            .MB_Title = "フォルダ作成確認"
            .MB_Button = Yes_No
            If Not .MBOX Then Exit Function
        End With
        'ﾃﾞｨﾚｸﾄﾘ作成
        CisFun.FE_Path = eKanban_MaeKoteiFld
        CisFun.DirCheck NoMsg_Crt
    End If
    
    OutFNo = FreeFile
    Open RTrim(eKanban_MaeKoteiFld) & RTrim(eKanban_MaeKoteiFile) For Output As #OutFNo

    Call CisFrm.MousePT(11)

    Proc_MSG = "前工程情報データ作成中"
    Proc_MSG.Visible = True: Bar.Visible = True: DoEvents

    With CisFrm
        Set .MeterPic = Bar
        .MeterMax = CisFun.Val2(lb_ChoiceSu)
    End With

' ( 出力処理 )
    Dim ID          As Long
    Dim Crt_Cnt     As Long
    Crt_Cnt = 0
    With vsGrid
        For ID = 2 To .Rows - 1
            If Trim(.TextMatrix(ID, CisVsGrid.FixedGet(1))) = "" Then Exit For
            If Trim(.TextMatrix(ID, CisVsGrid.FixedGet(18))) = "X" Then
                
                Call SQL_Crt(ID)
                
                If NTMRead(gSL_Select) Then
                                    
                   Do Until Not NTM_RDSTS
                      If Not TXT_Out Then
                         Call CisFrm.MousePT(1)
                         Exit Function
                      End If
                      
                      Call NTMReadNext
                   Loop
                    
                End If
                
                Call NTMClose
                                
                Crt_Cnt = Crt_Cnt + 1
                CisFrm.MeterUpdate Crt_Cnt: DoEvents
        
            End If
        Next ID
    End With

    On Error GoTo 0

    Proc_MSG.Visible = False
    
    Close #OutFNo
    
    TXTPut = True
    
    '----- 今回出力結果ﾌﾟﾙｰﾌﾘｽﾄ発行 -----
    '  確認メッセージ
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "    出力内容プルーフリスト発行を行います       "
        .MB_Title = "印刷確認"
        .MB_MSG(4) = "     発行を行いますか？        "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Call CisFrm.MousePT(1)
            Exit Function
        End If
    End With
    
    Call List_Out               'ﾌﾟﾙｰﾌﾘｽﾄ出力
    
    Call CisFrm.MousePT(1)

    Exit Function
TXTPut_Err:
    Proc_MSG.Visible = False
    
    Close #OutFNo
    
    Call CisFrm.MousePT(1)

    CisFun.ErrorBox
    End
End Function
Private Function TXT_Out() As Boolean

    TXT_Out = False
    
    Static sw_Str           As String

'-----     -----     -----     -----     < 2006/04/11 ADD >
    wMax = 1
    Erase wOki: Erase wSyuyo:   Erase wHako:    Erase wDate:    Erase wBin
    
    With NTM
        wOki(1) = RTrim$(.置場)
        wSyuyo(1) = .収容数1
        wHako(1) = RTrim$(.収容器1)
        wDate(1) = ""
        wBin(1) = ""

        If RTrim$(.変更日) <> "" Or RTrim$(.所番地変更日) <> "" Then
            
            If RTrim$(.変更日) <> "" And RTrim$(.所番地変更日) = "" Then
                .所番地変更日 = RTrim$(.変更日)
                .所番地変更便 = RTrim$(.変更便)
            End If
            
            If RTrim$(.変更日) = "" And RTrim$(.所番地変更日) <> "" Then
                .変更日 = RTrim$(.所番地変更日)
                .変更便 = RTrim$(.所番地変更便)
            End If
            
            If RTrim$(.変更日) & RTrim$(.変更便) = RTrim$(.所番地変更日) & RTrim$(.所番地変更便) Then
'
'                   [   変更日・便が置場・収容系共に同じか、一方のみが変更日・便指定されていた場合  ]
'
                wMax = 2
                If RTrim$(.変更後所番地) <> "" Then
                    wOki(2) = RTrim$(.変更後所番地)
                Else
                    wOki(2) = RTrim$(.置場)
                End If
                
                If .収容数2 <> 0 Then
                    wSyuyo(2) = .収容数2
                    wHako(2) = RTrim$(.収容器2)
                Else
                    wSyuyo(2) = .収容数1
                    wHako(2) = RTrim$(.収容器1)
                End If
                
                wDate(2) = RTrim$(.変更日)
                wBin(2) = RTrim$(.変更便)
            Else
'
'                   [   置場・収容系共に変更日・便指定があり、変更タイミングが異なる場合    ]
'
                wMax = 3
                If RTrim$(.変更日) & RTrim$(.変更便) < RTrim$(.所番地変更日) & RTrim$(.所番地変更便) Then
                    wOki(2) = RTrim$(.置場)
                    wSyuyo(2) = .収容数2
                    wHako(2) = RTrim$(.収容器2)
                    wDate(2) = RTrim$(.変更日)
                    wBin(2) = RTrim$(.変更便)
                
                    wOki(3) = RTrim$(.変更後所番地)
                    wSyuyo(3) = .収容数2
                    wHako(3) = RTrim$(.収容器2)
                    wDate(3) = RTrim$(.所番地変更日)
                    wBin(3) = RTrim$(.所番地変更便)
                Else
                    wOki(2) = RTrim$(.変更後所番地)
                    wSyuyo(2) = .収容数1
                    wHako(2) = RTrim$(.収容器1)
                    wDate(2) = RTrim$(.所番地変更日)
                    wBin(2) = RTrim$(.所番地変更便)
                
                    wOki(3) = RTrim$(.変更後所番地)
                    wSyuyo(3) = .収容数2
                    wHako(3) = RTrim$(.収容器2)
                    wDate(3) = RTrim$(.変更日)
                    wBin(3) = RTrim$(.変更便)
                End If
            End If
        End If
    End With
'-----     -----     -----     -----     < 2006/04/11 ADD END >
    
    For wCnt = 1 To wMax
        CrtVal = ""
        With NTM

            '1.設置場所コード(10桁)
            CrtVal = RTrim$(eKanban_SBasyoCD) & Space(10 - Len(RTrim$(eKanban_SBasyoCD))) & Chr(9)
            '2.マシン№(2桁)
            CrtVal = CrtVal & RTrim$(eKanban_MachineNo) & Space(2 - Len(RTrim$(eKanban_MachineNo))) & Chr(9)
            '3.納入先事業体(2桁)
            CrtVal = CrtVal & RTrim$(.納入先事業体) & Space(2 - Len(RTrim$(.納入先事業体))) & Chr(9)
            '4.納入先事業体工区(1桁)
            CrtVal = CrtVal & RTrim$(.納入先事業体工区) & Space(1 - Len(RTrim$(.納入先事業体工区))) & Chr(9)
            
'--  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  < 2006/07/04 MNT >
'                CrtVal = CrtVal & Left$(RTrim$(.取引先C) & Space(4), 4)

            '5.納入先(4桁)
            sw_Str = "0000" & RTrim$(.納入先コード)
            CrtVal = CrtVal & Mid$(sw_Str, Len(sw_Str) - 3, 4) & Chr(9)
            
'--  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  < 2006/07/04 MNT END >
            '6.納入先工区(1桁)
            CrtVal = CrtVal & RTrim$(.納入先工区) & Space(1 - Len(RTrim$(.納入先工区))) & Chr(9)
            '7.受入(2桁)
            CrtVal = CrtVal & RTrim$(.受入) & Space(2 - Len(RTrim$(.受入))) & Chr(9)
            '9.自工程(6桁)
            CrtVal = CrtVal & RTrim$(.自工程) & Space(6 - Len(RTrim$(.自工程))) & Chr(9)
            '10.仕入先事業体(2桁)
            CrtVal = CrtVal & RTrim$(.仕入先事業体) & Space(2 - Len(RTrim$(.仕入先事業体))) & Chr(9)
            '11.仕入先コード(4桁)
            CrtVal = CrtVal & RTrim$(.仕入先コード) & Space(4 - Len(RTrim$(.仕入先コード))) & Chr(9)
            '12.仕入先工区(1桁)
            CrtVal = CrtVal & RTrim$(.仕入先工区) & Space(1 - Len(RTrim$(.仕入先工区))) & Chr(9)
            '13.出荷場(3桁)
            CrtVal = CrtVal & RTrim$(.出荷場) & Space(3 - Len(RTrim$(.出荷場))) & Chr(9)
            '14.前工程(6桁)
            CrtVal = CrtVal & RTrim$(.前工程) & Space(6 - Len(RTrim$(.前工程))) & Chr(9)
            '15.品番(12桁)
            CrtVal = CrtVal & CisFun.Mid2(RTrim$(.品番), 1, 12) & Space(12 - Len(RTrim$(CisFun.Mid2(RTrim$(.品番), 1, 12)))) & Chr(9)
            '16.背番号(4桁)･･･右詰め
            CrtVal = CrtVal & Space(4 - Len(Trim$(.背番号))) & Trim$(.背番号) & Chr(9)
            
'-----     -----     -----     -----     < 2006/04/11 MNT >
'                CrtVal = CrtVal & Space(8)
'                CrtVal = CrtVal & Space(2)
            '17.切替オーダー№(YYYYMMDDnn形式,10桁)
            If wCnt = 1 Then
                CrtVal = CrtVal & "20000101"
                CrtVal = CrtVal & "01" & Chr(9)
            Else
                CrtVal = CrtVal & wDate(wCnt) & Space(8 - Len(wDate(wCnt)))
                CrtVal = CrtVal & wBin(wCnt) & Space(2 - Len(wBin(wCnt))) & Chr(9)
            End If
'-----     -----     -----     -----     < 2006/04/11 MNT END >
        
'       - 18.ユーザエリア -
            '表示品番(14桁)
            CrtVal = CrtVal & CisFun.Mid2(RTrim$(.表示品番), 1, 14) & Space(14 - Len(RTrim$(CisFun.Mid2(RTrim$(.表示品番), 1, 14))))
            
'           - 収容数・箱種 -
'-----     -----     -----     -----     -----     -----     -----     -----     -----     < 2006/04/11 MNT >
'                gLong = .収容数1
'                gStr = RTrim$(.収容器1)
'                If RTrim$(.変更日) <> "" Then
'                    If RTrim$(.変更日) <= Format(Now(), "YYYYMMDD") Then
'                        gLong = .収容数2
'                        gStr = RTrim$(.収容器2)
'                    End If
'                End If
'                CrtVal = CrtVal & RTrim$(gStr) & Space(10 - Len(RTrim$(gStr)))
            
'-----     -----     -----     -----     -----     -----     -----     -----     -----     < 2006/04/11 MNT END >
        
'           - 所番地 -
'-----     -----     -----     -----     -----     -----     -----     -----     -----     < 2006/04/11 MNT >
'                gStr = RTrim$(.置場)
'                If RTrim$(.所番地変更日) <> "" Then
'                    If RTrim$(.所番地変更日) <= Format(Now(), "YYYYMMDD") Then
'                        gStr = RTrim$(.変更後所番地)
'                    End If
'                End If
'                CrtVal = CrtVal & CisFun.Mid2(RTrim$(gStr), 1, 8) & Space(8 - Len(RTrim$(CisFun.Mid2(RTrim$(gStr), 1, 8))))
'                CrtVal = CrtVal & CisFun.Mid2(RTrim$(gStr), 1, 8) & Space(8 - Len(RTrim$(CisFun.Mid2(RTrim$(gStr), 1, 8))))
            
            '置場(空白2桁 + 8桁)
            CrtVal = CrtVal & Space(2) & CisFun.Mid2(RTrim$(wOki(wCnt)), 1, 8) & Space(8 - Len(RTrim$(CisFun.Mid2(RTrim$(wOki(wCnt)), 1, 8))))
'-----     -----     -----     -----     -----     -----     -----     -----     -----     < 2006/04/11 MNT END >
        
            '工場集荷コード(2桁)
            CrtVal = CrtVal & RTrim$(.工場集荷コード) & Space(2 - Len(RTrim$(.工場集荷コード)))
            
'           - 収容数 -
'-----     -----     -----     -----     -----     < 2006/04/11 MNT >
'                gStr = Format(gLong, "#")
            gStr = Format(wSyuyo(wCnt), "#")
'-----     -----     -----     -----     -----     < 2006/04/11 MNT END >
            
            If Len(gStr) > 5 Then
                CisFun.MB_Lines = 10
                CisFun.MB_MSG(1) = "　　　　　     【　　収容数桁数オーバー　　】      　　　　"
                CisFun.MB_MSG(3) = "　　　　　　　　品　番 : " & .表示品番
                CisFun.MB_MSG(4) = "　　　　　　　　収容数 : " & gStr
                CisFun.MB_MSG(6) = "　　処理を続行しますか？"
                CisFun.MB_MSG(7) = "　　　但し、当データの収容数は正しい値とは異なります      "
                
                CisFun.MB_Title = "桁数確認"
                CisFun.MB_Button = Yes_No
                If Not CisFun.MBOX Then
                    Call NTMClose
                    Close #OutFNo
                    Exit Function
                End If
                gStr = Left$(gStr, 5)
            End If
            
            '収容数(5桁)
            CrtVal = CrtVal & Space(5 - Len(Trim$(gStr))) & Trim$(gStr)
            
'-----     -----     -----     -----     -----     -----     < 2006/02/15 MNT START >
'
'            CrtVal = CrtVal & Space(26)
'           '工場集荷順路(3桁)
            gStr = Format(Format(.工場集荷順路, "#"), "@@@")
            CrtVal = CrtVal & gStr & Space(3 - Len(RTrim$(gStr)))
            'CrtVal = CrtVal & Format(Format(.工場集荷順路, "#"), "@@@")
            
            'FILLER(161桁)
            CrtVal = CrtVal & Space(161) & Chr(9)
'
'-----     -----     -----     -----     -----     -----     < 2006/02/15 MNT  END  >
            '19.重量原単位(5桁)･･･空白
            CrtVal = CrtVal & Space(5) & Chr(9)
            
            '20.*** 全角情報エリア ***
            gStr = Space(8)
            CrtVal = CrtVal & StrConv(gStr, vbWide)         'ブランク（全角）8桁
            
            gStr = RTrim$(.工場集荷コード) & Space(2 - Len(RTrim$(.工場集荷コード)))
            CrtVal = CrtVal & StrConv(gStr, vbWide)         '工場集荷コード（全角）2桁
            
            gStr = CisFun.Mid2(RTrim$(.工場集荷名称), 1, 16)
            gStr = StrConv(gStr, vbWide)
            gInt = CisFun.Len2(gStr)
            gInt = 16 - gInt
            gInt = gInt / 2
            
            gStr = gStr & Space(gInt)
            CrtVal = CrtVal & StrConv(gStr, vbWide)         '工場集荷名（全角）8桁
            
            '箱種(10桁)
            gStr = RTrim$(wHako(wCnt)) & Space(10 - Len(RTrim$(wHako(wCnt))))
            gStr = StrConv(gStr, vbWide)
            If CisFun.Len2(gStr) < 20 Then
               For gInt = CisFun.Len2(gStr) + 2 To 20 Step 2
                   gStr = gStr & "　"
               Next gInt
            End If
            
            CrtVal = CrtVal & gStr
            
            '製品マーク(4桁)････ﾎﾟｲﾝﾄ自体は3桁,4桁目は空白
            gStr = RTrim$(.ポイント) & Space(3 - Len(RTrim$(.ポイント))) & " "
            gStr = StrConv(gStr, vbWide)
            If CisFun.Len2(gStr) < 8 Then
               For gInt = CisFun.Len2(gStr) + 2 To 8 Step 2
                   gStr = gStr & "　"
               Next gInt
            End If
            
            CrtVal = CrtVal & gStr
            
            gStr = Space(18)
            CrtVal = CrtVal & StrConv(gStr, vbWide)         'ブランク（全角）18桁
        
        End With

        Print #OutFNo, CrtVal
    
    Next wCnt
        
    TXT_Out = True
    
End Function
Private Sub SQL_Crt(ID As Long)
    gSL_Select = "SELECT "
    gSL_Select = gSL_Select & " TR1.納入先事業体 , TR1.納入先事業体工区 , SUBSTRING(TR1.取引先,1,4) 納入先コード, SUBSTRING(TR1.工場CD,1,1) 納入先工区 , " & vbCrLf
    gSL_Select = gSL_Select & "HT.受入 , TR1.略称 納入先略称 , " & vbCrLf
    gSL_Select = gSL_Select & "TR2.仕入先事業体 , TR2.略称 契約先略称 , " & vbCrLf
    gSL_Select = gSL_Select & "SUBSTRING(TR2.被自社,1,4) 仕入先コード , SUBSTRING(TR2.被自社工場,1,1) 仕入先工区 , " & vbCrLf
    gSL_Select = gSL_Select & "ISNULL(ME.自工程,'') 自工程 , ISNULL(ME.前工程,'') 前工程 , " & vbCrLf
    gSL_Select = gSL_Select & "ISNULL(HT.出荷場 ,'') 出荷場 ,  " & vbCrLf
    
    gSL_Select = gSL_Select & "HT.品番 , ISNULl(HN.表示品番,'') 表示品番 , " & vbCrLf
    gSL_Select = gSL_Select & " HT.収容数1, HT.収容器1, HT.変更日, HT.変更便, HT.収容数2,HT.収容器2," & vbCrLf
    gSL_Select = gSL_Select & " HT.置場, HT.所番地変更日, HT.所番地変更便, HT.変更後所番地," & vbCrLf
    gSL_Select = gSL_Select & "SUBSTRING(HT.背番号,1,4) 背番号 , HT.工場集荷コード , " & vbCrLf
    gSL_Select = gSL_Select & "CASE WHEN ISNULL(HT.工場集荷コード,'') = '' THEN '' " & vbCrLf
    gSL_Select = gSL_Select & "ELSE ISNULL(SY.値名称,'') END 工場集荷名称 , " & vbCrLf
    gSL_Select = gSL_Select & "HN.ポイント , HT.工場集荷順路 " & vbCrLf
    
    gSL_Select = gSL_Select & "FROM 品番取引先マスタ HT " & vbCrLf
    gSL_Select = gSL_Select & "LEFT OUTER JOIN 品番マスタ HN " & vbCrLf
    gSL_Select = gSL_Select & "        ON  HN.品番 = HT.品番 " & vbCrLf
    gSL_Select = gSL_Select & "LEFT OUTER JOIN 取引先マスタ TR1 " & vbCrLf
    gSL_Select = gSL_Select & "        ON  HT.取引先 = TR1.取引先CD " & vbCrLf
    gSL_Select = gSL_Select & "        AND TR1.取引先区分 = 0 " & vbCrLf
    gSL_Select = gSL_Select & "LEFT OUTER JOIN 取引先マスタ TR2 " & vbCrLf
    gSL_Select = gSL_Select & "        ON  HT.契約先 = TR2.取引先CD " & vbCrLf
    gSL_Select = gSL_Select & "        AND TR2.取引先区分 = 0 " & vbCrLf
    gSL_Select = gSL_Select & "LEFT OUTER JOIN 新TOPPS前工程情報テーブル ME " & vbCrLf
    gSL_Select = gSL_Select & "        ON  ME.納入先 = HT.取引先 " & vbCrLf
    gSL_Select = gSL_Select & "        AND ME.受入 = HT.受入 " & vbCrLf
    gSL_Select = gSL_Select & "        AND ME.契約先 = HT.契約先 " & vbCrLf
    gSL_Select = gSL_Select & "LEFT OUTER JOIN 名称マスタ SY " & vbCrLf
    gSL_Select = gSL_Select & "        ON  SY.区分名称 = '工場集荷管理'" & vbCrLf
    gSL_Select = gSL_Select & "        AND SY.区分タイプ= '' " & vbCrLf
    gSL_Select = gSL_Select & "        AND SY.値 = ISNULL(HT.工場集荷コード,'')" & vbCrLf
    With CisFun
         .Where_CreateKB = NewDocument
            
         ' 納入先事業体
         .Where_Create AlphaMe, "ISNULL(TR1.納入先事業体,'')", RTrim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(12)))
         ' 納入先事業体工区
         .Where_Create AlphaMe, "ISNULL(TR1.納入先事業体工区,'')", RTrim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(13)))
         ' 納入先
         .Where_Create AlphaMe, "TR1.取引先", RTrim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(14)))
         ' 納入先工区
         .Where_Create AlphaMe, "TR1.工場CD", RTrim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(15)))
         ' 受入
         .Where_Create AlphaMe, "HT.受入", RTrim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(3)))
         ' 仕入先事業体
         .Where_Create AlphaMe, "ISNULL(TR2.仕入先事業体,'')", RTrim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(6)))
         ' 仕入先(契約先：被自社)
         .Where_Create AlphaMe, "SUBSTRING(ISNULL(TR2.被自社,''),1,4)", RTrim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(16)))
         ' 仕入先工区(契約先：被自社工場)
         .Where_Create AlphaMe, "SUBSTRING(ISNULL(TR2.被自社工場,''),1,1)", RTrim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(17)))
         
         gSL_Select = gSL_Select & .Where_Phrase
         
         ' 自工程
         If RTrim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(5))) <> "" Then
            gSL_Select = gSL_Select & " AND ISNULL(ME.自工程,'') = '" & RTrim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(5))) & "'"
         Else
            gSL_Select = gSL_Select & " AND ISNULL(ME.自工程,'') = ''"
         End If
         ' 出荷場
         If RTrim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(9))) <> "" Then
            gSL_Select = gSL_Select & " AND ISNULL(HT.出荷場,'') = '" & RTrim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(9))) & "'"
         Else
            gSL_Select = gSL_Select & " AND ISNULL(HT.出荷場,'') = ''"
         End If
         ' 前工程
         If RTrim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(10))) <> "" Then
            gSL_Select = gSL_Select & " AND ISNULL(ME.前工程,'') = '" & RTrim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(10))) & "'"
         Else
            gSL_Select = gSL_Select & " AND ISNULL(ME.前工程,'') = ''"
         End If
    
'         If RTrim(.Where_Phrase) <> "" Then
'            gSL_Select = gSL_Select & .Where_Phrase & " AND HT.品目 = 0 "
'         Else
'            gSL_Select = gSL_Select & " WHERE HT.品目 = 0 "
'         End If
    End With
    
    gSL_Select = gSL_Select & "AND HT.品目 = 0 " & vbCrLf
    gSL_Select = gSL_Select & "AND ISNULL(TR1.取引先CD,'') <> '' " & vbCrLf
    gSL_Select = gSL_Select & "AND ISNULL(TR2.取引先CD,'') <> '' " & vbCrLf
    gSL_Select = gSL_Select & "AND ISNULL(TR1.納入先事業体,'') <> '' " & vbCrLf
    gSL_Select = gSL_Select & "AND ISNULL(TR1.納入先事業体工区,'') <> '' " & vbCrLf
    gSL_Select = gSL_Select & "AND ISNULL(TR2.仕入先事業体,'') <> '' " & vbCrLf
    gSL_Select = gSL_Select & "AND ISNULL(TR2.被自社,'') <> '' " & vbCrLf
    
End Sub
'-----------------------------------------------------
'--  今回出力内容プルーフリスト作成
'-----------------------------------------------------
Private Sub List_Out()
    Dim ID          As Long

    With CisSimplePrint
        .PrintTitle = "※　前工程情報出力プルーフリスト　※"
        .Preservation = App.Title
        .PrintSize = A4
        .PrintOrientation = Landscape
        .PrintHeadLine = あり
        .PrintBodyLine = あり
        .PrintBodyGyo = 41
        
        
        .pHeadItemSet 0, 中, "納入先事業体"
        .pHeadItemSet 1, 中, "納入先ｺｰﾄﾞ"
        .pHeadItemSet 2, 中, "受入"
        .pHeadItemSet 3, 左, "納入先名称"
        .pHeadItemSet 4, 中, "自工程"
        
        .pHeadItemSet 5, 中, "仕入先事業体"
        .pHeadItemSet 6, 左, "仕入先事業体名(契約先名)  "
        .pHeadItemSet 7, 中, "仕入先ｺｰﾄﾞ"
        .pHeadItemSet 8, 中, "出荷場"
        .pHeadItemSet 9, 中, "前工程"
        
        .pHeadItemSet 10, 右, "品番件数"
    
        .pBodyItemSet 0, 12, 中
        .pBodyItemSet 1, 10, 中
        .pBodyItemSet 2, 4, 中
        .pBodyItemSet 3, 24, 左
        .pBodyItemSet 4, 6, 中
        
        .pBodyItemSet 5, 12, 中
        .pBodyItemSet 6, 24, 左
        .pBodyItemSet 7, 10, 中
        .pBodyItemSet 8, 6, 中
        .pBodyItemSet 9, 6, 中
        
        .pBodyItemSet 10, 8, 右
        
        .BodyInit
        .MaxCount = CisFun.Val2(lb_ChoiceSu)
        
        If Not .CheckPrtCsv Then Exit Sub
        
        For ID = 2 To vsGrid.Rows - 1
            If Trim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(1))) = "" Then Exit For
            If Trim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(18))) = "X" Then
                
                Call SQL_Crt(ID)
                
                If NTMRead(gSL_Select) Then
                   
                   .pBodyValueSet 0, Trim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(1)))
                   .pBodyValueSet 1, Trim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(2)))
                   .pBodyValueSet 2, Trim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(3)))
                   .pBodyValueSet 3, Trim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(4)))
                   .pBodyValueSet 4, Trim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(5)))
                   .pBodyValueSet 5, Trim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(6)))
                   .pBodyValueSet 6, Trim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(7)))
                   .pBodyValueSet 7, Trim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(8)))
                   .pBodyValueSet 8, Trim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(9)))
                   .pBodyValueSet 9, Trim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(10)))
                   .pBodyValueSet 10, Format(CisFun.Val2(Trim(vsGrid.TextMatrix(ID, CisVsGrid.FixedGet(11)))), "##,##0")
                    
                   .BodySet
                   
                End If
                
                Call NTMClose
                                
            End If
        Next ID
        
        .PrintStart
    
    End With
    

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
        Call .FixedSet(中中, 中中, あり, 2, 0, "")
        Call .FixedSet(中中, 中中, あり, 6, 1, "事業  ")
        Call .FixedSet(中中, 中中, あり, 10, 2, "納入先ｺｰﾄﾞ")
        Call .FixedSet(中中, 中中, あり, 4, 3, "受入")
        Call .FixedSet(中中, 左中, あり, 21, 4, "納入先名称")
        Call .FixedSet(中中, 中中, あり, 8, 5, "自工程")
        Call .FixedSet(中中, 中中, あり, 6, 6, "事業体")
        Call .FixedSet(中中, 左中, あり, 21, 7, "事業体名称(契約先)")
        Call .FixedSet(中中, 中中, なし, 10, 8, "仕入先ｺｰﾄﾞ")
        Call .FixedSet(中中, 中中, なし, 6, 9, "出荷場")
        Call .FixedSet(中中, 中中, なし, 8, 10, "前工程")
        Call .FixedSet(中中, 右中, なし, 4, 11, "件数")
        Call .FixedSet(中中, 左中, なし, 0, 12, "納入先事業体")
        Call .FixedSet(中中, 左中, なし, 0, 13, "納入先事業体工区")
        Call .FixedSet(中中, 左中, なし, 0, 14, "納入先")
        Call .FixedSet(中中, 左中, なし, 0, 15, "納入先工区")
        Call .FixedSet(中中, 左中, なし, 0, 16, "仕入先")
        Call .FixedSet(中中, 左中, なし, 0, 17, "仕入先工区")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 18, " Select")
'        Call .FixedSet(中中, 左中, なし, 0, 10, "Select")
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

    If NTMRead(gSL_Select, , 1) Then

'       抽出処理 初期ｾｯﾄ
        Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
'       ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
        With vsGrid
           .Redraw = False
            If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
                .Rows = CisVsGrid.DispMax
            Else
                .Rows = Ttl_Kensu + 2
            End If
        End With
        ID = 1
        
        lb_DispSu = Format(CisDB.RecordCount, "###,##0")      ' 抽出件数(合計)退避
        
        Do Until Not NTM_RDSTS
            ID = ID + 1

            Call NTMInfoToGrid(ID)

            Call NTMReadNext(1)
        Loop
        With vsGrid
             .Row = 2: .Col = 1: .ColSel = .Cols - 1
             .Redraw = True
        End With

        GridSet = True
    End If
    Call NTMClose
End Function
'+----------------------------------------------+
'+      テーブルよりグリッドに転送
'+----------------------------------------------+
Private Sub NTMInfoToGrid(SetRow As Long)
    With vsGrid
        '選択
        .TextMatrix(SetRow, 0) = ""
        '納入先事業体
        If RTrim(NTM.納入先事業体工区) <> "" Then
           .TextMatrix(SetRow, 1) = RTrim(NTM.納入先事業体) & "-" & RTrim(NTM.納入先事業体工区)
        Else
           .TextMatrix(SetRow, 1) = RTrim(NTM.納入先事業体)
        End If
        '納入先コード
        If RTrim(NTM.納入先工区) <> "" Then
           .TextMatrix(SetRow, 2) = RTrim(NTM.納入先コード) & "-" & RTrim(NTM.納入先工区)
        Else
           .TextMatrix(SetRow, 2) = RTrim(NTM.納入先コード)
        End If
        '受入
        .TextMatrix(SetRow, 3) = RTrim(NTM.受入)
        '納入先名称
        .TextMatrix(SetRow, 4) = RTrim(NTM.納入先略称)
        '自工程
        .TextMatrix(SetRow, 5) = RTrim(NTM.自工程)
        '仕入先事業体
        .TextMatrix(SetRow, 6) = RTrim(NTM.仕入先事業体)
        '契約先名称
        .TextMatrix(SetRow, 7) = RTrim(NTM.契約先略称)
        '仕入先コード
        If RTrim(NTM.仕入先工区) <> "" Then
           .TextMatrix(SetRow, 8) = RTrim(NTM.仕入先コード) & "-" & RTrim(NTM.仕入先工区)
        Else
           .TextMatrix(SetRow, 8) = RTrim(NTM.仕入先コード)
        End If
        '出荷場
        .TextMatrix(SetRow, 9) = RTrim(NTM.出荷場)
        '前工程
        .TextMatrix(SetRow, 10) = RTrim(NTM.前工程)
        '対象件数
        .TextMatrix(SetRow, 11) = Format(NTM.対象件数, "#,##0")
        '納入先事業体
        .TextMatrix(SetRow, 12) = RTrim(NTM.納入先事業体)
        '納入先事業体工区
        .TextMatrix(SetRow, 13) = RTrim(NTM.納入先事業体工区)
        '納入先ｺｰﾄﾞ
        .TextMatrix(SetRow, 14) = RTrim(NTM.納入先コード)
        '納入先工区
        .TextMatrix(SetRow, 15) = RTrim(NTM.納入先工区)
        '仕入先ｺｰﾄﾞ
        .TextMatrix(SetRow, 16) = RTrim(NTM.仕入先コード)
        '仕入先工区
        .TextMatrix(SetRow, 17) = RTrim(NTM.仕入先工区)
        'Select
        '.TextMatrix(SetRow, 18) = ""
    End With
End Sub
'******************************************'
'*****      行選択（クリック）
'******************************************'
Private Sub vsGrid_Click()
    If ProcHB = "B1" Then
        Call CisVsGrid.vsColor
    End If

'    If vsGrid.TextMatrix(vsGrid.Row, 18) = "" Then
'       lb_ChoiceSu = Format(CisFun.Val2(lb_ChoiceSu) + 1, "###,##0")
'    Else
'       lb_ChoiceSu = Format(CisFun.Val2(lb_ChoiceSu) - 1, "###,##0")
'    End If

End Sub
'******************************************'
'*****      行選択（スペース）
'******************************************'
Private Sub vsGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If ProcHB = "B1" Then
        If KeyCode = vbKeySpace Then
            Call CisVsGrid.vsColor
        
'            If vsGrid.TextMatrix(vsGrid.Row, 18) = "" Then
'               lb_ChoiceSu = Format(CisFun.Val2(lb_ChoiceSu) + 1, "###,##0")
'            Else
'               lb_ChoiceSu = Format(CisFun.Val2(lb_ChoiceSu) - 1, "###,##0")
'            End If
        
        End If
    End If
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

