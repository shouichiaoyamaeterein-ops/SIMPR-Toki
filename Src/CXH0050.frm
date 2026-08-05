VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXH0050 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "発注打切入力"
   ClientHeight    =   10575
   ClientLeft      =   225
   ClientTop       =   4800
   ClientWidth     =   14595
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
   ScaleHeight     =   10575
   ScaleWidth      =   14595
   WindowState     =   2  '最大化
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'なし
      Height          =   345
      Index           =   1
      Left            =   10575
      ScaleHeight     =   345
      ScaleWidth      =   3720
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   15
      Width           =   3720
      Begin Cis3D_v60.CIS3D CIS3D15 
         Height          =   345
         Left            =   2460
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         BackColor       =   8421631
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
         Left            =   30
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         BackColor       =   8454143
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
            Top             =   45
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
         Left            =   1245
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         BackColor       =   16777088
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
   Begin Cis3D_v60.CIS3D PB_UPD 
      Height          =   480
      Left            =   9930
      Top             =   1980
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   847
      BackColor       =   0
      ForeColor       =   16777215
      Caption         =   "品番一括作成"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      cAlingnment     =   7
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
      Left            =   9930
      ScaleHeight     =   1245
      ScaleWidth      =   4275
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   615
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
   Begin VB.PictureBox B3_Area1 
      Height          =   3930
      Left            =   1845
      ScaleHeight     =   3870
      ScaleWidth      =   9540
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6105
      Width           =   9600
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   450
         Left            =   0
         Top             =   15
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   794
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "<<  対象品番　表示  >>"
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
         cFont3D         =   1
         cAlingnment     =   7
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   345
            Left            =   7515
            Top             =   60
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   609
            Caption         =   " 表示件数"
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
            cAlingnment     =   1
            Begin Cis3D_v60.CIS3D B1lb_TCnt 
               Height          =   270
               Left            =   1065
               Top             =   30
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   476
               BackColor       =   12648384
               Caption         =   "#,###"
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
      End
      Begin vsFlexLib.vsFlexArray vsGrid1 
         Height          =   3390
         Left            =   60
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   480
         Width           =   9465
         _Version        =   196608
         _ExtentX        =   16695
         _ExtentY        =   5980
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
         FormatString    =   "№|仕入先|   仕  入  先  名   |受入|背番号 |品　番                   |構成|"
         Rows            =   11
         Cols            =   8
         BackColor       =   12648447
         BackColorBkg    =   14737632
         FocusRect       =   0
         HighLight       =   0
         RowHeightMin    =   300
         GridLines       =   2
         SelectionMode   =   1
         BackColorAlternate=   12648447
      End
   End
   Begin VB.PictureBox H1_Area2 
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
      Left            =   1875
      ScaleHeight     =   885
      ScaleWidth      =   7200
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2025
      Width           =   7260
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   780
         Left            =   60
         Top             =   60
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "品　　番"
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
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   360
            Left            =   135
            TabIndex        =   0
            Top             =   315
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
            IMEMode         =   2
            cDataReplace    =   1
            cFaZero         =   0
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   780
         Left            =   3480
         Top             =   60
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "取 引 先"
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
            TabIndex        =   1
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
      TabIndex        =   7
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
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 発注打切入力  】"
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
         Left            =   14325
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
         Left            =   7995
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
   Begin Cis3D_v60.CIS3D H1_Area1 
      Height          =   615
      Left            =   1890
      Top             =   1350
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   1085
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
         Height          =   510
         Left            =   60
         Top             =   60
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   900
         ForeColor       =   16711680
         Caption         =   "打切種別"
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
         cPositionX      =   50
         Begin VB.OptionButton OpH1_UKbn 
            Caption         =   "仕入先"
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
            Left            =   5355
            TabIndex        =   5
            Top             =   105
            Width           =   1050
         End
         Begin VB.OptionButton OpH1_UKbn 
            Caption         =   "品番・仕入先"
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
            Left            =   3195
            TabIndex        =   4
            Top             =   105
            Width           =   1890
         End
         Begin VB.OptionButton OpH1_UKbn 
            Caption         =   "品番"
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
            Left            =   2190
            TabIndex        =   3
            Top             =   105
            Width           =   855
         End
         Begin VB.OptionButton OpH1_UKbn 
            Caption         =   "残数"
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
            Index           =   0
            Left            =   1170
            TabIndex        =   2
            Top             =   105
            Value           =   -1  'True
            Width           =   885
         End
      End
   End
   Begin VB.PictureBox B2_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   1800
      Left            =   1845
      ScaleHeight     =   1740
      ScaleWidth      =   8025
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3480
      Width           =   8085
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   1620
         Index           =   1
         Left            =   75
         Top             =   75
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   2858
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
            Height          =   540
            Index           =   3
            Left            =   30
            Top             =   45
            Width           =   7830
            _ExtentX        =   13811
            _ExtentY        =   953
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
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   420
               Index           =   2
               Left            =   150
               Top             =   60
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   741
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cBoderStyle     =   1
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   4
               Left            =   570
               Top             =   60
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "打 切 日"
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
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   10
               Left            =   2400
               Top             =   60
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "警 告 日"
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
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   5
               Left            =   4275
               Top             =   60
               Width           =   3465
               _ExtentX        =   6112
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "切替仕入先"
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
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   495
            Index           =   11
            Left            =   30
            Top             =   1080
            Width           =   7830
            _ExtentX        =   13811
            _ExtentY        =   873
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
            Begin CisYMD_With_Btn.CisYMDwB iB2_UYmd 
               Height          =   375
               Left            =   555
               TabIndex        =   19
               Top             =   45
               Width           =   1830
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
               Object.Height          =   375
               Object.Width           =   1830
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
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   360
               Index           =   3
               Left            =   150
               Top             =   60
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   635
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "新"
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
            Begin CisText_V60.CisText iB2_KTorcd 
               Height          =   360
               Left            =   4290
               TabIndex        =   21
               Top             =   60
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
            Begin Cis3D_v60.CIS3D B2lb_KTornm 
               Height          =   360
               Left            =   5280
               Top             =   60
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
            Begin CisYMD_With_Btn.CisYMDwB iB2_KYmd 
               Height          =   375
               Left            =   2415
               TabIndex        =   20
               Top             =   45
               Width           =   1830
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
               Object.Height          =   375
               Object.Width           =   1830
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
            Begin CisText_V60.CisText iB2_KUkeir 
               Height          =   360
               Left            =   7755
               TabIndex        =   23
               Top             =   60
               Visible         =   0   'False
               Width           =   345
               _ExtentX        =   609
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
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   495
            Index           =   12
            Left            =   30
            Top             =   585
            Width           =   7830
            _ExtentX        =   13811
            _ExtentY        =   873
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
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   360
               Index           =   4
               Left            =   150
               Top             =   60
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   635
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "旧"
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
            Begin Cis3D_v60.CIS3D B2lb_UYmdK 
               Height          =   360
               Left            =   555
               Top             =   60
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   635
               BackColor       =   16777152
               ForeColor       =   0
               Caption         =   "9999年99月99日"
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
            Begin Cis3D_v60.CIS3D B2lb_KYmdK 
               Height          =   360
               Left            =   2415
               Top             =   60
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   635
               BackColor       =   16777152
               ForeColor       =   0
               Caption         =   "9999年99月99日"
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
            Begin Cis3D_v60.CIS3D B2lb_TorcdK 
               Height          =   360
               Left            =   4275
               Top             =   60
               Width           =   3465
               _ExtentX        =   6112
               _ExtentY        =   635
               BackColor       =   16777152
               Caption         =   "XXXXXXX:XXXXXXXXX1XXXXXXXXX2"
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
            Begin Cis3D_v60.CIS3D B2lb_UkeirK 
               Height          =   360
               Left            =   7770
               Top             =   60
               Visible         =   0   'False
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   635
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   30
            End
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   1800
      Left            =   1845
      ScaleHeight     =   1740
      ScaleWidth      =   10860
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3480
      Width           =   10920
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   1620
         Index           =   0
         Left            =   75
         Top             =   75
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   2858
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
            Height          =   540
            Index           =   67
            Left            =   30
            Top             =   45
            Width           =   10650
            _ExtentX        =   18785
            _ExtentY        =   953
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
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   420
               Index           =   68
               Left            =   150
               Top             =   60
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   741
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cBoderStyle     =   1
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   71
               Left            =   570
               Top             =   60
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "   起 算 日   -便"
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
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   77
               Left            =   4860
               Top             =   60
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "発注数"
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
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   80
               Left            =   2835
               Top             =   60
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "設定数"
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
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   83
               Left            =   3855
               Top             =   60
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "警告数"
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
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   86
               Left            =   5910
               Top             =   60
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "入荷数"
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
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   89
               Left            =   7050
               Top             =   60
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "設定残数"
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
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   420
               Index           =   1
               Left            =   8295
               Top             =   60
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   741
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "   完 了 日   -便"
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
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   495
            Index           =   0
            Left            =   30
            Top             =   1080
            Width           =   10650
            _ExtentX        =   18785
            _ExtentY        =   873
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
            Begin Cis3D_v60.CIS3D B1lb_HSuryo 
               Height          =   360
               Left            =   4860
               Top             =   75
               Width           =   1020
               _ExtentX        =   1799
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
               cPositionX      =   -30
            End
            Begin CisText_V60.CisText iB1_SetSu 
               Height          =   360
               Left            =   2820
               TabIndex        =   11
               Top             =   75
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D B1lb_NSuryo 
               Height          =   360
               Left            =   5910
               Top             =   75
               Width           =   1110
               _ExtentX        =   1958
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
               cPositionX      =   -30
            End
            Begin CisText_V60.CisText iB1_KeiSu 
               Height          =   360
               Left            =   3840
               TabIndex        =   12
               Top             =   75
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D B1lb_UZanSu 
               Height          =   360
               Left            =   7050
               Top             =   75
               Width           =   1215
               _ExtentX        =   2143
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
               cPositionX      =   -30
            End
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   360
               Index           =   1
               Left            =   150
               Top             =   60
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   635
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "新"
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
            Begin CisText_V60.CisText iB1_KisBin 
               Height          =   375
               Left            =   2400
               TabIndex        =   9
               Top             =   60
               Width           =   390
               _ExtentX        =   688
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
            Begin CisText_V60.CisText iB1_KanBin 
               Height          =   360
               Left            =   10140
               TabIndex        =   15
               Top             =   75
               Width           =   390
               _ExtentX        =   688
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
            Begin CisYMD_With_Btn.CisYMDwB iB1_KisYmd 
               Height          =   360
               Left            =   555
               TabIndex        =   8
               Top             =   60
               Width           =   1830
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
               Object.Width           =   1830
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
            Begin CisYMD_With_Btn.CisYMDwB iB1_KanYmd 
               Height          =   345
               Left            =   8295
               TabIndex        =   13
               Top             =   75
               Width           =   1830
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
               Object.Width           =   1830
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
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   495
            Index           =   2
            Left            =   30
            Top             =   585
            Width           =   10650
            _ExtentX        =   18785
            _ExtentY        =   873
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
            Begin Cis3D_v60.CIS3D B1lb_HSuryoK 
               Height          =   360
               Left            =   4860
               Top             =   75
               Width           =   1020
               _ExtentX        =   1799
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
               cPositionX      =   -30
            End
            Begin Cis3D_v60.CIS3D B1lb_NSuryoK 
               Height          =   360
               Left            =   5910
               Top             =   75
               Width           =   1110
               _ExtentX        =   1958
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
               cPositionX      =   -30
            End
            Begin Cis3D_v60.CIS3D B1lb_UZanSuK 
               Height          =   360
               Left            =   7050
               Top             =   75
               Width           =   1215
               _ExtentX        =   2143
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
               cPositionX      =   -30
            End
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   360
               Index           =   0
               Left            =   150
               Top             =   60
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   635
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "旧"
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
            Begin Cis3D_v60.CIS3D B1lb_KisYmdK 
               Height          =   360
               Left            =   555
               Top             =   60
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   635
               BackColor       =   16777152
               ForeColor       =   0
               Caption         =   "9999年99月99日"
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
            Begin Cis3D_v60.CIS3D B1lb_KisBinK 
               Height          =   360
               Left            =   2400
               Top             =   60
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   635
               BackColor       =   16777152
               ForeColor       =   0
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
               cBoderStyle     =   1
            End
            Begin Cis3D_v60.CIS3D B1lb_SetSuK 
               Height          =   360
               Left            =   2820
               Top             =   75
               Width           =   990
               _ExtentX        =   1746
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
               cPositionX      =   -30
            End
            Begin Cis3D_v60.CIS3D B1lb_KeiSuK 
               Height          =   360
               Left            =   3840
               Top             =   75
               Width           =   990
               _ExtentX        =   1746
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
               cPositionX      =   -30
            End
            Begin Cis3D_v60.CIS3D B1lb_KanYmdK 
               Height          =   360
               Left            =   8295
               Top             =   75
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   635
               BackColor       =   16777152
               ForeColor       =   0
               Caption         =   "9999年99月99日"
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
            Begin Cis3D_v60.CIS3D B1lb_KanBinK 
               Height          =   360
               Left            =   10140
               Top             =   75
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   635
               BackColor       =   16777152
               ForeColor       =   0
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
               cBoderStyle     =   1
            End
         End
      End
   End
End
Attribute VB_Name = "CXH0050"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   発注打切入力
'**       フォームID    :   CXH0050
'**       処理概要      :
'**
'**       作  成  日    :   2003/12/19  By CIS SIMPR-A
'**       変  更  日    :   2008/04/11  By CIS 背番号８桁対応
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid1      As New CisVsGrid3
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim wPrintOut       As Integer      ' 出力先

    Dim MenuCall        As Boolean      ' 起動元
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   <><><> 起動元の判定 <><><>
    If RV_Call = "" Then
        MenuCall = True
    Else
        MenuCall = False
    End If

    If MenuCall Then
    '   #------------------#
    '   # 二 重 起 動 防 止 #
    '   #------------------#
        If Not Dupli_Start Then End
    '   #------------------#
    '   # 初 期 内 容 取 得 #
    '   #------------------#
        If Not IniGet Then End
    End If
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
    With CisVsGrid1
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 300
        .PatanMax = 1
        .SelectCol = 8
        .InitGet = False
        .Init
    End With
    Call GridInit1
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    If MenuCall Then
        SyoriKB = "ADD"
    Else
        SyoriKB = "MNT"
    End If
    Call SyoriName(SyoriKB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    If MenuCall Then
        Me.Show
    End If
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'初期表示
    If MenuCall Then
        Call OpH1_UKbn_Click(0)
    Else
        Select Case HUT.打切種別
            Case 0: Call OpH1_UKbn_Click(0)
            Case 1: Call OpH1_UKbn_Click(1)
            Case 2: Call OpH1_UKbn_Click(2)
            Case 3: Call OpH1_UKbn_Click(3)
        End Select
        iH1_Hinbn = HUT.品番
        iH1_Torcd = HUT.仕入先
        Call ReturnPress
    End If
'
    FormAct = Not FormAct
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
           Case vbKeyF5:       If PB_UPD.Visible Then Call PB_UPD_Click    '【品番一括作成】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyF8:       If PB_PRT.Visible Then Call PB_PRT_Click    '【印刷】
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  '【前進】
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  '【後退】
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
    If MenuCall Then
       Cancel = 1
    End If
    Call PB_END_Click
End Sub
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    If MenuCall Then
       '+---------------------+
       '+   ﾃﾞｰﾀﾍﾞｰｽ切断
       '+---------------------+
        Call CisDB.DBDISConnect
       '+---------------------+
       '+   ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
       '+---------------------+
        End
    Else
        Unload CXH0050
    End If
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
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKB)
      GoTo PB_CAN_Ed
   End If
   If ProcHB = "B2" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKB)
      PB_UPD.Visible = False
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
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If SyoriKB <> "MNT" Then
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
    End If
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
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = " SELECT HU.*,TR.略称 仕入先名 FROM 発注打切テーブル HU"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON  TR.取引先区分 = 1 "
    gSL_Select = gSL_Select & "   AND TR.取引先CD = HU.切替仕入先 "
    gSL_Select = gSL_Select & " WHERE CONVERT(VARCHAR(1),打切種別)"
    gSL_Select = gSL_Select & "       + 品番 + SPACE(25-LEN(品番))"
    gSL_Select = gSL_Select & "       + 仕入先 + SPACE(7-LEN(仕入先))"
    gSL_Select = gSL_Select & " < '" & RTrim(OpH1_UKbn(0).Tag) & RTrim(iH1_Hinbn) & Space(25 - Len(iH1_Hinbn))
    gSL_Select = gSL_Select & RTrim(iH1_Torcd) & Space(7 - Len(iH1_Torcd)) & "'"
    gSL_Select = gSL_Select & " ORDER BY 打切種別 desc,品番 desc,仕入先 Desc "
    Call DBInput("Back")
    Dummy.Enabled = True: Dummy.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = " SELECT HU.*,TR.略称 仕入先名 FROM 発注打切テーブル HU"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON  TR.取引先区分 = 1 "
    gSL_Select = gSL_Select & "   AND TR.取引先CD = HU.切替仕入先 "
    gSL_Select = gSL_Select & " WHERE CONVERT(VARCHAR(1),打切種別)"
    gSL_Select = gSL_Select & "       + 品番 + SPACE(25-LEN(品番))"
    gSL_Select = gSL_Select & "       + 仕入先 + SPACE(7-LEN(仕入先))"
    gSL_Select = gSL_Select & " > '" & RTrim(OpH1_UKbn(0).Tag) & RTrim(iH1_Hinbn) & Space(25 - Len(iH1_Hinbn))
    gSL_Select = gSL_Select & RTrim(iH1_Torcd) & Space(7 - Len(iH1_Torcd)) & "'"
    gSL_Select = gSL_Select & " ORDER BY 打切種別,品番,仕入先 "
    Call DBInput("Next")
    Dummy.Enabled = True: Dummy.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    印　刷  キ ー(F8) +
'+----------------------------+
Private Sub PB_PRT_Click()
    For gInt = 1 To 3
        If OpH1_UKbn(gInt) Then cUKbnm = OpH1_UKbn(gInt).Caption
    Next gInt
    cUYmd = B2lb_UYmdK
    cKYmd = B2lb_KYmdK
'    wPrintOut = 0               ' 出力先　：0=Printer Else 1=Preview
    
'リスト出力先確認
    With CisSimplePrint
         .PrintTitle = "打切対象品リスト"
         .SaveKey = gRegKey
         .Preservation = "CXH00050"
         If Not .CheckPrtCsv Then Exit Sub
    
         Select Case .PrintStatus
            Case 4: wPrintOut = 0
            Case 5: wPrintOut = 1
            Case 3: wPrintOut = 2
            Case Else: Exit Sub
         End Select
         
    End With
'印刷処理
    Call CXH0050LPrint(OpH1_UKbn(0).Tag, RTrim(iH1_Hinbn), RTrim(iH1_Torcd), wPrintOut)
    Call CisSimplePrint.CheckFormUnload
End Sub
'+----------------------------+
'+    品番一括作成      (F5) +
'+----------------------------+
Private Sub PB_UPD_Click()
    Call Hinbn_Create
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd(0): Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iB1_KisYmd" Then iB1_KisYmd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iB1_KanYmd" Then
       iB1_KanYmd.ShowCalender
       Exit Sub
    End If
    If PB_Look.Tag = "iB2_UYmd" Then
       iB2_UYmd.ShowCalender
       Exit Sub
    End If
    If PB_Look.Tag = "iB2_KYmd" Then
       iB2_KYmd.ShowCalender
       Exit Sub
    End If
    If PB_Look.Tag = "iB2_KTorcd" Then Call Look_Torcd(1): Exit Sub
'    If PB_Look.Tag = "iB2_KUkeir" Then Call Look_Ukeir: Exit Sub
End Sub
'+----------------------------+
'+       取引先検索            +
'+----------------------------+
Private Sub Look_Torcd(HBKbn As Byte)

    RV_Left = 0:    RV_Top = 0
    RV_TorKb = 8: RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        If HBKbn = 0 Then
            iH1_Torcd = RTrim(RV_TorcdK)
            Call iH1_Torcd_LostFocus
        Else
            iB2_KTorcd = RTrim(RV_TorcdK)
            Call iB2_KTorcd_LostFocus
        End If
    End If

End Sub
'+----------------------------+
'+       取引先・受入　検索            +
'+----------------------------+
Private Sub Look_Ukeir()

    RV_Left = 0:    RV_Top = 0
    RV_TorKb = 8: RV_Call = "C"
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iB2_KTorcd = RTrim(RV_TorcdK)
        Call iB2_KTorcd_LostFocus
'        iB2_KUkeir = RTrim(RV_Ukeir)
    End If

End Sub
'+----------------------------+
'+       品番検索            +
'+----------------------------+
Private Sub Look_Hinbn()

    RV_Left = 0: RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 1
    RVI_SK = 9
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing
'
    If RV_Rtn Then
        If OpH1_UKbn(0).Tag = "2" Then
            iH1_Torcd = RTrim(RV_TorcdK)
            Call iH1_Torcd_LostFocus
        End If
        iH1_Hinbn = RTrim(RV_Hinbn)
    End If

End Sub
'+-----------------------------+
'+  OpH1_UKbn   打切種別
'+-----------------------------+
Private Sub OpH1_UKbn_Click(INDEX As Integer)
    OpH1_UKbn(0).BackColor = gPLostSel
    OpH1_UKbn(1).BackColor = gPLostSel
    OpH1_UKbn(2).BackColor = gPLostSel
    OpH1_UKbn(3).BackColor = gPLostSel
    OpH1_UKbn(INDEX).BackColor = &HC0FFC0
    OpH1_UKbn(0).Tag = Format(INDEX, "0")
    If INDEX = 3 Then
        iH1_Hinbn = ""
        iH1_Hinbn.Enabled = False
    Else
        iH1_Hinbn.Enabled = True
    End If
    If INDEX = 0 Or INDEX = 1 Then
        iH1_Torcd = "": H1lb_Tornm = ""
        iH1_Torcd.Enabled = False
    Else
        iH1_Torcd.Enabled = True
    End If
    If INDEX = 0 Then
        B1_Area1.Visible = True
        B2_Area1.Visible = False
    Else
        B1_Area1.Visible = False
        B2_Area1.Visible = True
    End If
End Sub
'+-----------------------------+
'+  iH1_Hinbn
'+-----------------------------+
Private Sub iH1_Hinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
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
'+  iB1_KisYmd
'+-----------------------------+
Private Sub iB1_KisYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_KisYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    Call HSuryo_Set
    Call NSuryo_Set(0)
End Sub
'+-----------------------------+
'+  iB1_KisYmd
'+-----------------------------+
Private Sub iB1_KisBin_LostFocus()
    Call HSuryo_Set
    Call NSuryo_Set(0)
End Sub
'+-----------------------------+
'+  iB1_SetSu
'+-----------------------------+
Private Sub iB1_SetSu_LostFocus()
    B1lb_UZanSu = Format(CisFun.Val2(iB1_SetSu) - CisFun.Val2(B1lb_HSuryo), "#,###")
End Sub
'+-----------------------------+
'+  iB1_KanYmd
'+-----------------------------+
Private Sub iB1_KanYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_KanYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB2_UYmd
'+-----------------------------+
Private Sub iB2_UYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB2_UYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB2_KYmd
'+-----------------------------+
Private Sub iB2_KYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB2_KYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB2_KTorcd
'+-----------------------------+
Private Sub iB2_KTorcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB2_KTorcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B2lb_KTornm = ""
    If RTrim(iB2_KTorcd) <> "" Then
        If TorNmGet(iB2_KTorcd, 8) Then
            B2lb_KTornm = RTrim(TRM.略称)
        End If
    End If
End Sub
''''+-----------------------------+
''''+  iB2_KUkeir
''''+-----------------------------+
'''Private Sub iB2_KUkeir_GotFocus()
'''    PB_Look.Tag = ActiveControl.Name
'''    PB_Look.Visible = True
'''End Sub
'''Private Sub iB2_iB2_KUkeir_LostFocus()
'''    If ActiveControl.Name <> "PB_Look" Then
'''        PB_Look.Tag = ""
'''        PB_Look.Visible = False
'''    End If
'''End Sub
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

    H1_Area1.Enabled = H1Mode:  H1_Area1.BackColor = H1Color
    H1_Area2.Enabled = H1Mode:  H1_Area2.BackColor = H1Color
    B1_Area1.Enabled = B1Mode:  B1_Area1.BackColor = B1Color
    B2_Area1.Enabled = B2Mode:  B2_Area1.BackColor = B2Color

    PB_Look.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_ADD.Visible = H1Mode
    PB_MNT.Visible = H1Mode
    PB_DEL.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    PB_PRT.Visible = B2Mode
    If RTrim(B2lb_TorcdK) <> "" Then
        PB_UPD.Visible = True
    Else
        PB_UPD.Visible = False
    End If

    If MenuCall Then
        PB_END.BtnCaption = "終了"
    Else
        PB_END.BtnCaption = "戻る"
        PB_CAN.Visible = False
        PB_END.Visible = True
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
        CisVsGrid1.Clear
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
            If OpH1_UKbn(0).Tag = "0" Then
                ProcHB = "B1"
            Else
                ProcHB = "B2"
            End If
            Call DispChange(ProcHB)
        End If
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Or ProcHB = "B2" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        If MenuCall Then
            Call HeadBodyClear("B")
            ProcHB = "H1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
        Else
            Call PB_END_Click
        End If
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
'打切種別が　残数、品番、品番・仕入先の時
    If OpH1_UKbn(0).Tag <> "3" Then
    '品番
        If Trim(iH1_Hinbn) = "" Then
            CisFun.MB_MSG(2) = "    品番を入力して下さい            "
            CisFun.MBOX
            iH1_Hinbn.SetFocus
            Exit Function
        End If
        HNM.品番 = RTrim(iH1_Hinbn)
        If Not HNMRead("", 1) Then
            CisFun.MB_MSG(2) = "    品番　未登録！ （品番マスタ）          "
            CisFun.MBOX
            iH1_Hinbn.SetFocus
            Exit Function
        End If
        gSL_Select = "SELECT COUNT(品番) 件数 FROM 品番取引先マスタ"
        gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & "   AND 品目 = 1"
        If HTMRead(gSL_Select, 1) Then
            If HTM.件数 = 0 Then
                CisFun.MB_MSG(2) = "    品番　未登録！ （品番取引先マスタ）          "
                CisFun.MBOX
                iH1_Hinbn.SetFocus
                Exit Function
            End If
        End If
    End If
'打切種別が　品番・仕入先、仕入先の時
    If OpH1_UKbn(0).Tag = "2" Or OpH1_UKbn(0).Tag = "3" Then
    '仕入先
        If Trim(iH1_Torcd) = "" Then
            CisFun.MB_MSG(2) = "    仕入先を入力して下さい            "
            CisFun.MBOX
            iH1_Torcd.SetFocus
            Exit Function
        End If
        If Not TorNmGet(iH1_Torcd, 8) Then
            CisFun.MB_MSG(2) = "    仕入先　未登録！（取引先マスタ）            "
            CisFun.MBOX
            iH1_Torcd.SetFocus
            Exit Function
        End If
        gSL_Select = "SELECT COUNT(品番) 件数 FROM 品番取引先マスタ"
        gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   AND 品目 = 1"
        If OpH1_UKbn(0).Tag = "2" Then
            gSL_Select = gSL_Select & " AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
        End If
        If HTMRead(gSL_Select, 1) Then
            If HTM.件数 = 0 Then
                CisFun.MB_MSG(2) = "    品番　未登録！ （品番取引先マスタ）          "
                CisFun.MBOX
                iH1_Torcd.SetFocus
                Exit Function
            End If
        End If
    End If
' マスタ入力 SELECT文 ｾｯﾄ
    gSL_Select = " SELECT HU.*,TR.略称 仕入先名 FROM 発注打切テーブル HU"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON  TR.取引先区分 = 1 "
    gSL_Select = gSL_Select & "   AND TR.取引先CD = HU.切替仕入先 "
    gSL_Select = gSL_Select & " WHERE HU.打切種別 = " & CisFun.Val2(OpH1_UKbn(0).Tag) & ""
    If OpH1_UKbn(0).Tag <> "3" Then
        gSL_Select = gSL_Select & " AND HU.品番 = '" & RTrim(iH1_Hinbn) & "'"
        If OpH1_UKbn(0).Tag = "2" Then
            gSL_Select = gSL_Select & " AND HU.仕入先 = '" & RTrim(iH1_Torcd) & "'"
        End If
    Else
        gSL_Select = gSL_Select & " AND HU.仕入先 = '" & RTrim(iH1_Torcd) & "'"
    End If
    If DBInput("INV") Then
        If SyoriKB = "ADD" Then
            CisFun.MB_MSG(2) = "    打切情報　登録済！           "
            CisFun.MBOX
            Exit Function
        End If
    Else
        If SyoriKB <> "ADD" Then
            CisFun.MB_MSG(2) = "    打切情報　未登録！           "
            CisFun.MBOX
            Exit Function
        End If
    End If
    If OpH1_UKbn(0).Tag = "3" Then
        iB2_KTorcd.Enabled = True
'''        iB2_KUkeir.Enabled = True
    Else
        iB2_KTorcd.Enabled = False
'''        iB2_KUkeir.Enabled = False
        If SyoriKB = "ADD" Then
            gSL_Select = "SELECT COUNT(品番) 件数 FROM 発注打切テーブル"
            gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
            gSL_Select = gSL_Select & " AND 打切種別 <> " & CisFun.Val2(OpH1_UKbn(0).Tag) & ""
            gSL_Select = gSL_Select & " AND 打切種別 <> 3"
            If HUTRead(gSL_Select, 1) Then
                If HUT.件数 <> 0 Then
                    CisFun.MB_MSG(2) = "    この品番は他の種別で既に登録されています！           "
                    CisFun.MBOX
                    Exit Function
                End If
            End If
        End If
    End If
'Grid表示(対象品番を品番取引先マスタよりGridに表示）
    gSL_Select = " SELECT HT.*,TR.略称 取引先名,HM.表示品番"
    gSL_Select = gSL_Select & ",("
    gSL_Select = gSL_Select & "select count(*) from 構成マスタ KO"
    gSL_Select = gSL_Select & " where KO.親品番 = HT.品番"
    gSL_Select = gSL_Select & "    or KO.子品番 = HT.品番"
    gSL_Select = gSL_Select & " ) 件数"
    gSL_Select = gSL_Select & " FROM 品番取引先マスタ HT"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON  TR.取引先区分 = 1 "
    gSL_Select = gSL_Select & "   AND TR.取引先CD = HT.取引先 "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HM"
    gSL_Select = gSL_Select & "   ON HM.品番 = HT.品番 "
    If OpH1_UKbn(0).Tag <> "3" Then
        gSL_Select = gSL_Select & " WHERE HT.品番 = '" & RTrim(iH1_Hinbn) & "'"
        If OpH1_UKbn(0).Tag = "2" Then
            gSL_Select = gSL_Select & " AND HT.取引先 = '" & RTrim(iH1_Torcd) & "'"
        End If
    Else
        gSL_Select = gSL_Select & " WHERE HT.取引先 = '" & RTrim(iH1_Torcd) & "'"
    End If
    gSL_Select = gSL_Select & " AND HT.品目 = 1"
    If HTMRead(gSL_Select, 0) Then
        Call GridSet
    End If
    Call HTMClose
'
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
    With CisFun
        .MB_Lines = 4
        .MB_Title = "再入力"
        .MB_Button = Error
    End With
    
    If OpH1_UKbn(0).Tag = "0" Then
    '<< 打切種別=残数のチェック >>
    '起算日
        If RTrim(iB1_KisYmd) = "" Then
            CisFun.MB_MSG(2) = "    起算日　必須！            "
            CisFun.MBOX
            iB1_KisYmd.SetFocus
            Exit Function
        End If
        If Not iB1_KisYmd.cChkResult Then
            CisFun.MB_MSG(2) = "    起算日　異常！            "
            CisFun.MBOX
            iB1_KisYmd.SetFocus
            Exit Function
        End If
    '設定数
        If CisFun.Val2(iB1_SetSu) = 0 Then
            CisFun.MB_MSG(2) = "    設定数　必須！            "
            CisFun.MBOX
            iB1_SetSu.SetFocus
            Exit Function
        End If
    '警告数
        If CisFun.Val2(iB1_KeiSu) <> 0 Then
            If CisFun.Val2(iB1_KeiSu) > CisFun.Val2(iB1_SetSu) Then
                CisFun.MB_MSG(2) = "    警告数 > 設定数 ！            "
                CisFun.MBOX
                iB1_KeiSu.SetFocus
                Exit Function
            End If
        End If
    '設定残数
        If CisFun.Val2(B1lb_UZanSu) <= 0 Then
            CisFun.MB_MSG(2) = "    打切完了の更新となりますがよろしいですか？           "
            CisFun.MB_Button = OK_CAN
            If Not CisFun.MBOX Then
                iB1_SetSu.SetFocus
                Exit Function
            End If
        End If
    '完了日
        If CisFun.Val2(B1lb_UZanSu) <= 0 And RTrim(iB1_KanYmd) = "" Then
            CisFun.MB_MSG(2) = "    打切完了の更新時、完了日　必須！         "
            CisFun.MBOX
            iB1_KanYmd.SetFocus
            Exit Function
        End If
        If RTrim(iB1_KanYmd) <> "" And Not iB1_KanYmd.cChkResult Then
            CisFun.MB_MSG(2) = "    完了日　異常！            "
            CisFun.MBOX
            iB1_KanYmd.SetFocus
            Exit Function
        End If
        If RTrim(iB1_KanYmd) <> "" And RTrim(iB1_KanYmd) > Format(Now(), "yyyymmdd") Then
            CisFun.MB_MSG(2) = "    完了日 > システム日付！            "
            CisFun.MBOX
            iB1_KanYmd.SetFocus
            Exit Function
        End If
    Else
    '<< 打切種別=品番、品番･仕入先、仕入先のチェック >>
    '打切日
        If RTrim(iB2_UYmd) = "" Then
            CisFun.MB_MSG(2) = "    打切日　必須！            "
            CisFun.MBOX
            iB2_UYmd.SetFocus
            Exit Function
        End If
        If Not iB2_UYmd.cChkResult Then
            CisFun.MB_MSG(2) = "    打切日　異常！            "
            CisFun.MBOX
            iB2_UYmd.SetFocus
            Exit Function
        End If
    '警告日
        If RTrim(iB2_KYmd) <> "" And Not iB2_KYmd.cChkResult Then
            CisFun.MB_MSG(2) = "    警告日　異常！            "
            CisFun.MBOX
            iB2_KYmd.SetFocus
            Exit Function
        End If
        If RTrim(iB2_KYmd) <> "" And RTrim(iB2_KYmd) >= RTrim(iB2_UYmd) Then
            CisFun.MB_MSG(2) = "    警告日 > システム日付！            "
            CisFun.MBOX
            iB2_KYmd.SetFocus
            Exit Function
        End If
    '切替仕入先
        If RTrim(iB2_KTorcd) <> "" Then
            If Not TorNmGet(iB2_KTorcd, 8) Then
                CisFun.MB_MSG(2) = "    切替仕入先　未登録！（取引先マスタ）        "
                CisFun.MBOX
                iB2_KTorcd.SetFocus
                Exit Function
            End If
'''            If RTrim(iB2_KUkeir) <> "" Then
'''                TUM.取引先CD = RTrim(iB2_KTorcd)
'''                TUM.受入 = RTrim(iB2_KUkeir)
'''                TUM.取引先区分 = 1
'''                If Not TUMRead("", 1) Then
'''                    CisFun.MB_MSG(2) = "    切替受入　未登録！（取引先受入マスタ）        "
'''                    CisFun.MBOX
'''                    iB2_KUkeir.SetFocus
'''                    Exit Function
'''                End If
'''            End If
        End If
    End If
    
    
    Body1Chk = True
    Exit Function
Body1Chk_Err:
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
    
    If Not HUTRead(gSL_Select, 1) Then
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
'
    With HUT
        OpH1_UKbn(.打切種別) = True
        OpH1_UKbn(.打切種別).Tag = Format(.打切種別, "0")
        iH1_Hinbn = RTrim(.品番)
        iH1_Torcd = RTrim(.仕入先)
        Call iH1_Torcd_LostFocus
        
        If .打切種別 = 0 Then
            B1lb_KisYmdK = Left(.起算日, 4) & "年" & Mid(.起算日, 5, 2) & "月" & Mid(.起算日, 7, 2) & "日"
            B1lb_KisYmdK.Tag = RTrim(.起算日)
            B1lb_KisBinK = RTrim(.起算便)
            B1lb_KisBinK.Tag = RTrim(.起算便)
            B1lb_SetSuK = Format(.設定数, "#,###")
            B1lb_KeiSuK = Format(.警告数, "#,###")
            B1lb_HSuryoK = Format(.発注数, "#,###")
            Call NSuryo_Set(1)
            B1lb_UZanSuK = Format(.設定残数, "#,###")
            If RTrim(.完了日) <> "" Then
                B1lb_KanYmdK = Left(.完了日, 4) & "年" & Mid(.完了日, 5, 2) & "月" & Mid(.完了日, 7, 2) & "日"
                B1lb_KanBinK = RTrim(.完了便)
            End If
        Else
            B2lb_UYmdK = Left(.打切日, 4) & "年" & Mid(.打切日, 5, 2) & "月" & Mid(.打切日, 7, 2) & "日"
            B2lb_UYmdK.Tag = RTrim(.打切日)
            B2lb_KYmdK = Left(.警告日, 4) & "年" & Mid(.警告日, 5, 2) & "月" & Mid(.警告日, 7, 2) & "日"
            If RTrim(.切替仕入先) <> "" Then
                B2lb_TorcdK = RTrim(.切替仕入先) & ":" & RTrim(.仕入先名)
            End If
            B2lb_TorcdK.Tag = RTrim(.切替仕入先)
'''            B2lb_UkeirK = RTrim(.切替受入)
            
        End If
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
        ' 発注サイクルマスタ
        .SQL = "DELETE FROM 発注打切テーブル "
        .SQL = .SQL & " WHERE 打切種別 = " & CisFun.Val2(OpH1_UKbn(0).Tag) & ""
        .SQL = .SQL & "   And 品番 = '" & RTrim(iH1_Hinbn) & "'"
        .SQL = .SQL & "   And 仕入先 = '" & RTrim(iH1_Torcd) & "'"
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
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

    Call HSuryo_Set
'   マスタ項目移送
    Call ItemsClearHUT
    With HUT
        .打切種別 = CisFun.Val2(OpH1_UKbn(0).Tag)
        .品番 = RTrim(iH1_Hinbn)
        .仕入先 = RTrim(iH1_Torcd)
        If .打切種別 = 0 Then
            .起算日 = RTrim(iB1_KisYmd)
            .起算便 = Format(CisFun.Val2(iB1_KisBin), "00")
            .設定数 = CisFun.Val2(iB1_SetSu)
            .警告数 = CisFun.Val2(iB1_KeiSu)
            .発注数 = CisFun.Val2(B1lb_HSuryo)
            .入荷数 = CisFun.Val2(B1lb_NSuryo)
            .設定残数 = .設定数 - .発注数
            .完了日 = RTrim(iB1_KanYmd)
            .完了便 = RTrim(iB1_KanBin)
        Else
            .打切日 = RTrim(iB2_UYmd)
            .警告日 = RTrim(iB2_KYmd)
            .切替仕入先 = RTrim(iB2_KTorcd)
'''            .切替受入 = RTrim(iB2_KUkeir)
        End If
    End With
'
    If SyoriKB = "ADD" Then
        HUT.作成者 = gTanto
    ' ( 追加処理 )
        Call HUTInsert
    Else
        HUT.更新者 = gTanto
    ' ( 更新処理 )
        Call HUTUpdate
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
'****************************
'*      発注数　表示
'****************************
Private Function HSuryo_Set() As Boolean
    HSuryo_Set = False

    B1lb_HSuryo = ""
'SELECT文 ｾｯﾄ
    gSL_Select = " SELECT SUM(発注数)発注数 FROM 発注テーブル "
    gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " AND 納入日 + 納入便 >= '" & RTrim(iB1_KisYmd) & RTrim(iB1_KisBin) & "'"
    gSL_Select = gSL_Select & " UNION "
    gSL_Select = gSL_Select & " SELECT SUM(発注数)発注数 FROM 発注かんばん予定テーブル "
    gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " AND 納入日 + 納入便 >= '" & RTrim(iB1_KisYmd) & RTrim(iB1_KisBin) & "'"
    gSL_Select = gSL_Select & " And 処理区分 <> 2"
    If HCTRead(gSL_Select, 0) Then
        Do Until Not HCT_RDSTS
            If HCT.発注数 <> 0 Then
                B1lb_HSuryo = Format(CisFun.Val2(B1lb_HSuryo) + HCT.発注数, "#,###")
            End If
            Call HCTReadNext
         Loop
    End If
    Call HCTClose
    B1lb_UZanSu = Format(CisFun.Val2(iB1_SetSu) - CisFun.Val2(B1lb_HSuryo), "#,###")
'
    HSuryo_Set = True
End Function
'****************************
'*      入荷数　表示
'****************************
Private Function NSuryo_Set(KZKbn As Byte) As Boolean
    NSuryo_Set = False

    B1lb_NSuryo = ""
'SELECT文 ｾｯﾄ
    gSL_Select = " SELECT SUM(入荷数)入荷数 FROM 入荷テーブル "
    gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " AND 入荷日 + 入荷便 >= '"
    If KZKbn = 0 Then
        gSL_Select = gSL_Select & RTrim(iB1_KisYmd) & RTrim(iB1_KisBin) & "'"
    Else
        gSL_Select = gSL_Select & RTrim(B1lb_KisYmdK.Tag) & RTrim(B1lb_KisBinK.Tag) & "'"
    End If
    If NYTRead(gSL_Select, 1) Then
        If NYT.入荷数 <> 0 Then
            If KZKbn = 0 Then
                B1lb_NSuryo = Format(NYT.入荷数, "#,###")
            Else
                B1lb_NSuryoK = Format(NYT.入荷数, "#,###")
            End If
        End If
    End If
'
    NSuryo_Set = True
End Function
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit1()
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設
'定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 左中, あり, 7, 1, "仕入先")
        Call .FixedSet(中中, 左中, あり, 20, 2, "仕入先名")
        Call .FixedSet(中中, 中中, あり, 4, 3, "受入")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
        'Call .FixedSet(中中, 左中, あり, 6, 4, "背番号")
        Call .FixedSet(中中, 左中, あり, 8, 4, "背番号")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
        Call .FixedSet(中中, 左中, あり, 25, 5, "品　番")
        Call .FixedSet(中中, 中中, あり, 4, 6, "構成")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 7, "Select")
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
    B1lb_TCnt = Format(Ttl_Kensu, "#,###")
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid1
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid1.FixedGet(CisFun.Val2(ID))
        Next ID
        If Ttl_Kensu + 1 > .Rows Then
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not HTM_RDSTS
        'Grid Row,Col
        ID = ID + 1
        With vsGrid1
            .TextMatrix(ID, GridCol(0)) = ID
            .TextMatrix(ID, GridCol(1)) = RTrim(HTM.取引先)
            .TextMatrix(ID, GridCol(2)) = RTrim(HTM.取引先名)
            .TextMatrix(ID, GridCol(3)) = RTrim(HTM.受入)
            .TextMatrix(ID, GridCol(4)) = RTrim(HTM.背番号)
            .TextMatrix(ID, GridCol(5)) = RTrim(HTM.表示品番)
            If HTM.件数 = 0 Then
                .TextMatrix(ID, GridCol(6)) = ""
            Else
                .TextMatrix(ID, GridCol(6)) = "○"
            End If
        End With
        Call HTMReadNext
    Loop

    With vsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'+----------------------+
'+    品番一括作成       +
'+----------------------+
Private Function Hinbn_Create()
    Hinbn_Create = False
'取引先チェック
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
    If Not TorNmGet(B2lb_TorcdK.Tag, 9) Then
        CisFun.MB_MSG(2) = "    切替仕入先　未登録！（取引先マスタ）            "
        CisFun.MBOX
        iB2_KTorcd.SetFocus
        Exit Function
    End If
'確認
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 6
        .MB_MSG(1) = "   仕入先 = " & iH1_Torcd & ":" & RTrim(H1lb_Tornm) & " の品番(" & B1lb_TCnt & "件)を　　"
        .MB_MSG(2) = "   仕入先 = " & B2lb_TorcdK & "  で　　"
        .MB_MSG(3) = "   品番マスタに複写登録します。  "
        .MB_MSG(5) = "     よろしいですか？      "
        .MB_Title = "処理開始確認"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
'品番取引先マスタを作成
    gSL_Select = "INSERT INTO 品番取引先マスタ"
    gSL_Select = gSL_Select & " (品番,取引先,受入,アドレス,品目,背番号,SK区分,か係数,手配区分,"
    gSL_Select = gSL_Select & "  納入先,納入先受入,契約先,サイクル,収容数1,収容器1,変更日,"
    gSL_Select = gSL_Select & " 変更便,収容数2,収容器2,置場,最終発行連番,"
    gSL_Select = gSL_Select & " 伝票種類,納品書形態,受領書形態,受領書発行,明細書発行,"
    gSL_Select = gSL_Select & " 平準化区分,平準化係数,係数1,係数2,係数3,LT,ロット1,ロット2,"
    gSL_Select = gSL_Select & " Week1,Week2,Week3,Week4,Week5,Day1,Day2,Day3,Day4,Day5,"
    gSL_Select = gSL_Select & " 内示発注,発注比率,有効期限開始,有効期限終了,基準工数,PNO,備考,"
    gSL_Select = gSL_Select & " 作成日,作成者,作成端末,更新日,更新者,更新端末)"
    gSL_Select = gSL_Select & " SELECT HT.品番,'" & RTrim(B2lb_TorcdK.Tag) & "',HT.受入,'',HT.品目,HT.背番号,HT.SK区分,"
    gSL_Select = gSL_Select & "     HT.か係数,HT.手配区分,HT.納入先,HT.納入先受入,TR.契約先,"
    gSL_Select = gSL_Select & "     CASE WHEN ISNULL(TU.サイクル,'') = '' THEN HT.サイクル ELSE TU.サイクル END,"
    gSL_Select = gSL_Select & "     HT.収容数1,HT.収容器1,HT.変更日,HT.変更便,HT.収容数2,HT.収容器2,HT.置場,"
    gSL_Select = gSL_Select & "     0,TR.伝票種類,TR.納品書形態,TR.受領書形態,TR.受領書発行,TR.明細書発行,"
    gSL_Select = gSL_Select & "     TR.平準化区分,TR.平準化係数,HT.係数1,HT.係数2,HT.係数3,HT.LT,HT.ロット1,HT.ロット2,"
    gSL_Select = gSL_Select & "     HT.Week1,HT.Week2,HT.Week3,HT.Week4,HT.Week5,HT.Day1,HT.Day2,HT.Day3,HT.Day4,HT.Day5,"
    gSL_Select = gSL_Select & "     HT.内示発注,0,'" & RTrim(B2lb_UYmdK.Tag) & "','',HT.基準工数,HT.PNO,HT.備考,"
    If RTrim(gTanto) = "" Then
        gSL_Select = gSL_Select & "     GETDATE(),'CXH0050',HOST_NAME(),NULL,'',''"
    Else
        gSL_Select = gSL_Select & "     GETDATE(),'" & RTrim(gTanto) & "',HOST_NAME(),NULL,'',''"
    End If
    gSL_Select = gSL_Select & "  FROM 品番取引先マスタ HT"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HK"
    gSL_Select = gSL_Select & "    ON   HK.取引先 = '" & RTrim(B2lb_TorcdK.Tag) & "'"
    gSL_Select = gSL_Select & "    AND HK.品番 = HT.品番"
    gSL_Select = gSL_Select & "    AND HK.品目 = 1"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "    ON  TR.取引先CD = '" & RTrim(B2lb_TorcdK.Tag) & "'"
    gSL_Select = gSL_Select & "    AND TR.取引先区分 = 1"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先受入マスタ TU"
    gSL_Select = gSL_Select & "    ON   TU.取引先CD = '" & RTrim(B2lb_TorcdK.Tag) & "'"
    gSL_Select = gSL_Select & "    AND TU.受入 = HT.受入"
    gSL_Select = gSL_Select & "    AND TU.取引先区分 = 1"
    gSL_Select = gSL_Select & " WHERE HT.取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "    AND HT.品目 = 1"
    gSL_Select = gSL_Select & "    AND HK.品番 IS NULL"
    With CisDB
        On Error GoTo Hinbn_Create_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
        .SQL = gSL_Select
        .DBExec
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        On Error GoTo 0
    End With
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "   仕入先 = " & B2lb_TorcdK & "  の品番を (" & CisDB.RecordCount & " 件)作成しました。"
        .MB_Title = "結果確認"
        .MB_Button = OK
        .MBOX
    End With
    
    Hinbn_Create = True
    Exit Function
Hinbn_Create_Err:
    CisFun.ErrorBox
    End
End Function

