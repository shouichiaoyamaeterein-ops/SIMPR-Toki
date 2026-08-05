VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CXM1400 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "工程かんばん発行内容パターン登録"
   ClientHeight    =   11235
   ClientLeft      =   1275
   ClientTop       =   6315
   ClientWidth     =   15360
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Cis3D_v60.CIS3D CIS3D12 
      Height          =   300
      Left            =   8400
      Top             =   3480
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   529
      BackColor       =   -2147483633
      Caption         =   "返却場所１"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   -2147483633
      cBoderColor2    =   -2147483633
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cFont3D         =   2
      cBoderStyle     =   2
   End
   Begin Cis3D_v60.CIS3D CIS3D11 
      Height          =   300
      Left            =   3780
      Top             =   3480
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   529
      BackColor       =   -2147483633
      Caption         =   "部署１"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   -2147483633
      cBoderColor2    =   -2147483633
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cFont3D         =   2
      cBoderStyle     =   2
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
      Left            =   8520
      ScaleHeight     =   1260
      ScaleWidth      =   4320
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1680
      Width           =   4380
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1755
         Top             =   645
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
         Left            =   2610
         Top             =   645
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
         Left            =   3465
         Top             =   645
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
      Begin CisBtn_60.CisBtn PB_BACK 
         Height          =   585
         Left            =   2610
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
      Begin CisBtn_60.CisBtn PB_NEXT 
         Height          =   585
         Left            =   3465
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   1755
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
         Left            =   30
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
         Left            =   900
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   900
         Top             =   645
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
   Begin VB.PictureBox H1_Area1 
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
      Left            =   2760
      ScaleHeight     =   1260
      ScaleWidth      =   5460
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1680
      Width           =   5520
      Begin Cis3D_v60.CIS3D UC_3D 
         Height          =   615
         Left            =   30
         Top             =   30
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   1085
         ForeColor       =   16711680
         Caption         =   "部　　署"
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
         cPositionX      =   120
         cPositionY      =   15
         Begin CisText_V60.CisText iH1_BusyoCD 
            Height          =   375
            Left            =   1170
            TabIndex        =   5
            Top             =   105
            Width           =   960
            _ExtentX        =   1693
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
            Text            =   "XXXXXXX"
            MaxLength       =   7
            cAutoEject      =   0   'False
         End
         Begin Cis3D_v60.CIS3D H1lb_BusyoN 
            Height          =   360
            Left            =   2130
            Top             =   105
            Width           =   2505
            _ExtentX        =   4419
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
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   600
         Left            =   30
         Top             =   645
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   1058
         ForeColor       =   16711680
         Caption         =   "返却場所"
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
         cPositionX      =   120
         Begin CisText_V60.CisText iH1_BasyoCD 
            Height          =   375
            Left            =   1170
            TabIndex        =   6
            Top             =   90
            Width           =   465
            _ExtentX        =   820
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
            cAutoEject      =   0   'False
         End
         Begin Cis3D_v60.CIS3D H1lb_BasyoN 
            Height          =   360
            Left            =   1635
            Top             =   90
            Width           =   3690
            _ExtentX        =   6509
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
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
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   3885
      Left            =   2790
      ScaleHeight     =   3825
      ScaleWidth      =   10050
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3330
      Width           =   10110
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   3690
         Index           =   1
         Left            =   75
         Top             =   75
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   6509
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
            Height          =   5010
            Left            =   -120
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   0
            Width           =   10140
            _Version        =   196608
            _ExtentX        =   17886
            _ExtentY        =   8837
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
            FormatString    =   "登録№|部署   |１                            |返 |１                            |x部|x返|Select||"
            Rows            =   11
            BackColor       =   12648447
            BackColorBkg    =   12632256
            FocusRect       =   2
            HighLight       =   2
            RowHeightMin    =   330
            SelectionMode   =   1
            Begin Cis3D_v60.CIS3D CIS3D2 
               Height          =   330
               Left            =   9570
               Top             =   0
               Visible         =   0   'False
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   582
               BackColor       =   -2147483633
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
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cFont3D         =   2
            End
            Begin Cis3D_v60.CIS3D Input_Row 
               Height          =   315
               Left            =   795
               Top             =   330
               Width           =   9090
               _ExtentX        =   16034
               _ExtentY        =   556
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
               Begin CisText_V60.CisText iB1_BusyoCD 
                  Height          =   330
                  Left            =   0
                  TabIndex        =   8
                  Top             =   0
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  cGotColor       =   12640511
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
               Begin Cis3D_v60.CIS3D B1lb_BusyoN 
                  Height          =   330
                  Left            =   945
                  Top             =   0
                  Width           =   3675
                  _ExtentX        =   6482
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
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
                  cFont3D         =   2
                  cAlingnment     =   1
                  cBoderStyle     =   1
                  cPositionX      =   30
               End
               Begin CisText_V60.CisText iB1_BasyoCD 
                  Height          =   330
                  Left            =   4620
                  TabIndex        =   9
                  Top             =   0
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  cGotColor       =   12640511
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
               Begin Cis3D_v60.CIS3D B1lb_BasyoN 
                  Height          =   330
                  Left            =   5100
                  Top             =   0
                  Width           =   3675
                  _ExtentX        =   6482
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
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
                  cFont3D         =   2
                  cAlingnment     =   1
                  cBoderStyle     =   1
                  cPositionX      =   30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D1 
               Height          =   3645
               Left            =   9570
               Top             =   0
               Visible         =   0   'False
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   6429
               BackColor       =   -2147483633
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
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cFont3D         =   2
            End
         End
         Begin CisText_V60.CisText B_TabDummy 
            Height          =   285
            Left            =   5685
            TabIndex        =   2
            Top             =   60
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
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
      Left            =   240
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   400
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Left            =   900
      Top             =   300
   End
   Begin VB.Timer Timer2 
      Left            =   1170
      Top             =   300
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
      Caption         =   "【 工程かんばん発行内容パターン登録 】"
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
   Begin Cis3D_v60.CIS3D Kensu1 
      Height          =   360
      Left            =   12000
      Top             =   1080
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   635
      BackColor       =   16761024
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
      cFont3DColor1   =   8421504
      cFont3DColor2   =   -2147483643
      cFont3D         =   2
      cAlingnment     =   1
      cBoderStyle     =   1
      cPositionX      =   30
   End
   Begin Cis3D_v60.CIS3D Kensu 
      Height          =   360
      Left            =   10440
      Top             =   1080
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   635
      BackColor       =   16761024
      Caption         =   "削除選択件数"
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
   Begin Cis3D_v60.CIS3D Kensu2 
      Height          =   360
      Left            =   12000
      Top             =   720
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   635
      BackColor       =   16761024
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
      cFont3DColor1   =   8421504
      cFont3DColor2   =   -2147483643
      cFont3D         =   2
      cAlingnment     =   1
      cBoderStyle     =   1
      cPositionX      =   30
   End
   Begin Cis3D_v60.CIS3D CIS3D3 
      Height          =   360
      Left            =   10440
      Top             =   720
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   635
      BackColor       =   16761024
      Caption         =   "修正対象件数"
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
Attribute VB_Name = "CXM1400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   工程かんばん発行内容パターン登録
'**       フォームID    :   CXM1400
'**       処理概要      :
'**
'**       作  成  日    :   2009/10/22  By CIS SIMPR-A
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim ActObj          As Object
    Dim ErrObj          As Object
    Dim ActRow          As Integer
    
    Dim RD_Count        As Long
    Dim Dsp_Flg         As Byte

    Dim mBusyoCD()      As String
    Dim mBasyoCD()      As String
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
        .NumberSet = No     ' NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 7
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
    SyoriKB = ""
    Call SyoriName(SyoriKB)





B1_Area1.Height = 4500
UC_3D4(1).Height = 4440







'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)
    Dsp_Flg = 0
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
'   >> 名称マスタ確認
    Call MeisyoInit

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
           Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click     '【新規】
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click     '【修正】
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click     '【削除】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click   '【検索】
           Case vbKeyPageUp:
                            If ProcHB = "H1" Then
                                If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
                            Else
                                Exit Function
                            End If
           Case vbKeyPageDown:
           
                            If ProcHB = "H1" Then
                                If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
                            Else
                                Exit Function
                            End If
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click     '【取消】
                                
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click     '【終了】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                       '【入力】
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
Private Sub PB_CAN_Click()
   Call CisFrm.MousePT(11)
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
   
   If ProcHB = "H1" Then
      SyoriKB = ""
      Call SyoriName(SyoriKB)
      PB_CAN.Visible = False                    ' 2009.10.22 add
      iH1_BusyoCD.SetFocus
      GoTo PB_CAN_Ed
   End If
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKB)
      iH1_BusyoCD.SetFocus
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+------------------------------+
'+    新　規  キ ー(F1)         +
'+------------------------------+
Private Sub PB_ADD_Click()

    If SyoriKB <> "ADD" Then
       SyoriKB = "ADD"
       PB_CAN.Visible = True
    Else
       SyoriKB = ""
       PB_CAN.Visible = False
    End If
    Call SyoriName(SyoriKB)
End Sub
'+------------------------------+
'+    修　正  キ ー(F2)         +
'+------------------------------+
Private Sub PB_MNT_Click()

    If SyoriKB <> "MNT" Then
       SyoriKB = "MNT"
       PB_CAN.Visible = True
    Else
       SyoriKB = ""
       PB_CAN.Visible = False
    End If
    Call SyoriName(SyoriKB)
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)    +
'+----------------------------+
Private Sub PB_DEL_Click()
   
    If SyoriKB <> "DEL" Then     ' 2009.10.22 upd : = "" → <> "DEL" に変更
        SyoriKB = "DEL"
        PB_CAN.Visible = True
    Else
        SyoriKB = ""
        PB_CAN.Visible = False
    End If
    Call SyoriName(SyoriKB)
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = "SELECT M1.*"
    gSL_Select = gSL_Select & " FROM  名称マスタ M1"
    gSL_Select = gSL_Select & " WHERE M1.区分名称 = '工程かんばん発行パターン'"
    gSL_Select = gSL_Select & "   AND SUBSTRING(M1.値名称 + '       ', 1, 7) + M1.英数字1 < "
    gSL_Select = gSL_Select & "       '" & Mid(Trim(iH1_BusyoCD) & Space(7), 1, 7) & _
                                           Mid(Trim(iH1_BasyoCD) & Space(3), 1, 3) & "'"       ' 部署 7桁 + 返却場所 3桁"
'    gSL_Select = gSL_Select & "   AND RTRIM(M1.値名称) <> ""                "
    gSL_Select = gSL_Select & " ORDER BY 値名称 DESC, 英数字1 DESC"

    If Not SYMRead(gSL_Select, 1) Or Trim(SYM.値名称) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     先頭レコードです            "
            .MB_Title = "【後退】"
            .MB_Button = OK
            .MBOX
        End With
        GoTo PB_Back_Ed
    End If
    iH1_BusyoCD = Trim(SYM.値名称):  Call iH1_BusyoCD_LostFocus
    iH1_BasyoCD = Trim(SYM.英数字1): Call iH1_BasyoCD_LostFocus
    
    gSL_Select = "SELECT M1.*, T2.略称 AS 名称名, M3.値名称 AS 検索名"
    gSL_Select = gSL_Select & " FROM  名称マスタ M1"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ T2"
    gSL_Select = gSL_Select & "            ON  T2.取引先CD   = M1.英数字2"
    gSL_Select = gSL_Select & "            AND T2.取引先区分 = 1"           ' 取引先区分：発注
'   gSL_Select = gSL_Select & "            AND T2.社内区分   = 1"           ' 社内区分　：社内
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ   M3"
    gSL_Select = gSL_Select & "            ON  M3.区分名称   = '返却場所'"
    gSL_Select = gSL_Select & "            AND M3.値         = M1.英数字3"
    gSL_Select = gSL_Select & " WHERE M1.区分名称 = '工程かんばん発行パターン'"
    gSL_Select = gSL_Select & "   AND SUBSTRING(M1.値名称 + '       ', 1, 7) + M1.英数字1 = "
    gSL_Select = gSL_Select & "       '" & Mid(Trim(iH1_BusyoCD) & Space(7), 1, 7) & _
                                           Mid(Trim(iH1_BasyoCD) & Space(3), 1, 3) & "'"       ' 部署 7桁 + 返却場所 3桁"
    gSL_Select = gSL_Select & " ORDER BY 値 , 値名称, 英数字1"
    
    Call DBInput("Back")
PB_Back_Ed:
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_BusyoCD.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub

'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)

    gSL_Select = "SELECT M1.*"
    gSL_Select = gSL_Select & " FROM  名称マスタ M1"
    gSL_Select = gSL_Select & " WHERE M1.区分名称 = '工程かんばん発行パターン'"
    gSL_Select = gSL_Select & "   AND SUBSTRING(M1.値名称 + '       ', 1, 7) + M1.英数字1 > "
    gSL_Select = gSL_Select & "       '" & Mid(Trim(iH1_BusyoCD) & Space(7), 1, 7) & _
                                           Mid(Trim(iH1_BasyoCD) & Space(3), 1, 3) & "'"       ' 部署 7桁 + 返却場所 3桁"
    gSL_Select = gSL_Select & " ORDER BY 値名称, 英数字1"
    If Not SYMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     最終レコードです            "
            .MB_Title = "【前進】"
            .MB_Button = OK
            .MBOX
        End With
        GoTo PB_Next_Ed
    End If
    iH1_BusyoCD = Trim(SYM.値名称):  Call iH1_BusyoCD_LostFocus
    iH1_BasyoCD = Trim(SYM.英数字1): Call iH1_BasyoCD_LostFocus
    
    gSL_Select = "SELECT M1.*, T2.略称 AS 名称名, M3.値名称 AS 検索名"
    gSL_Select = gSL_Select & " FROM  名称マスタ M1"
   'gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ T1"
   'gSL_Select = gSL_Select & "            ON  T1.取引先CD = M1.値名称"
   'gSL_Select = gSL_Select & "            AND T1.社内区分 = 1"
   'gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ   M2"
   'gSL_Select = gSL_Select & "            ON  M2.区分名称 = '返却場所'"
   'gSL_Select = gSL_Select & "            AND M2.値       = M1.英数字1"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ T2"
    gSL_Select = gSL_Select & "            ON  T2.取引先CD   = M1.英数字2"
    gSL_Select = gSL_Select & "            AND T2.取引先区分 = 1"           ' 取引先区分：発注
'   gSL_Select = gSL_Select & "            AND T2.社内区分   = 1"           ' 社内区分　：社内
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ   M3"
    gSL_Select = gSL_Select & "            ON  M3.区分名称   = '返却場所'"
    gSL_Select = gSL_Select & "            AND M3.値         = M1.英数字3"
    gSL_Select = gSL_Select & " WHERE M1.区分名称 = '工程かんばん発行パターン'"
    gSL_Select = gSL_Select & "   AND SUBSTRING(M1.値名称 + '       ', 1, 7) + M1.英数字1 = "
    gSL_Select = gSL_Select & "       '" & Mid(Trim(iH1_BusyoCD) & Space(7), 1, 7) & _
                                           Mid(Trim(iH1_BasyoCD) & Space(3), 1, 3) & "'"       ' 部署 7桁 + 返却場所 3桁"
    gSL_Select = gSL_Select & " ORDER BY 値, 値名称, 英数字1"
    Call DBInput("Next")
PB_Next_Ed:
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_BusyoCD.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_BusyoCD" Then Call Look_Busyo(0): Exit Sub
    If PB_Look.Tag = "iH1_BasyoCD" Then Call Look_Basyo(0): Exit Sub
    If PB_Look.Tag = "iB1_BusyoCD" Then Call Look_Busyo(1): Exit Sub
    If PB_Look.Tag = "iB1_BasyoCD" Then Call Look_Basyo(1): Exit Sub
End Sub
'+----------------------------+
'+       部署検索             +
'+----------------------------+
Private Sub Look_Busyo(KB As Integer)

    RV_Left = 0
    RV_Top = 0
    If KB = 0 Then
        RV_TorKb = 5
    Else
        RV_TorKb = 8
    End If
    RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        If KB = 0 Then
            iH1_BusyoCD = RTrim(RV_TorcdK)
            H1lb_BusyoN = RV_TorRName
           'Call iH1_BusyoCD_LostFocus
        Else
            iB1_BusyoCD = RTrim(RV_TorcdK)
            B1lb_BusyoN = RV_TorRName
           'Call iB1_BusyoCD_LostFocus
        End If
    End If

End Sub
'+----------------------------+
'+       名称マスタ検索        +
'+       (返却場所検索)       +
'+----------------------------+
Private Sub Look_Basyo(KB As Integer)

    RV_Left = 0
    RV_Top = 0
    RVI_Kubun = "返却場所"
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
       Select Case KB
            Case 0
                iH1_BasyoCD = RV_Code
                H1lb_BasyoN = RV_Kubun
               'Call iH1_BasyoCD_LostFocus
            Case 1
                iB1_BasyoCD = RV_Code
                B1lb_BasyoN = RV_Kubun
               'Call iB1_BasyoCD_LostFocus
       End Select
    End If

End Sub

'+-----------------------------+
'+  iH1_BusyoCD　部署
'+-----------------------------+
Private Sub iH1_BusyoCD_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_BusyoCD_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H1lb_BusyoN = ""
    If RTrim(iH1_BusyoCD) <> "" Then
       If TorNmGet(iH1_BusyoCD, 9) Then
          H1lb_BusyoN = TRM_RName       ' TRM.略称
       End If
    End If
End Sub
'+-----------------------------+
'+  iH1_BasyoCD　返却場所
'+-----------------------------+
Private Sub iH1_BasyoCD_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_BasyoCD_LostFocus()
    Dim wSel    As String
    
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = True
    End If
    
    H1lb_BasyoN = ""
    If RTrim(iH1_BasyoCD) <> "" Then
       wSel = "SELECT * FROM 名称マスタ"
       wSel = wSel & " WHERE 区分名称 = '返却場所'"
       wSel = wSel & "   AND 値       = '" & Trim(iH1_BasyoCD.Text) & "'"

       If SYMRead(wSel) = True Then
          H1lb_BasyoN = SYM.値名称
       End If
    End If
End Sub
'+-----------------------------+
'+  iB1_BusyoCD　部署１
'+-----------------------------+
Private Sub iB1_BusyoCD_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_BusyoCD_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_BusyoCD_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B1lb_BusyoN = ""
    If RTrim(iB1_BusyoCD) <> "" Then
       If TorNmGet(iB1_BusyoCD, 9) Then
          B1lb_BusyoN = TRM_RName       ' RTrim(TRM.略称)
       End If
    End If
'    Call iB1_Chk
'    Call iB1_BasyoCD.SetFocus
End Sub
'+-----------------------------+
'+  iH1_BasyoCD　返却場所１
'+-----------------------------+
Private Sub iB1_BasyoCD_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_BasyoCD_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_BasyoCD_LostFocus()
    Dim wSel    As String
    
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = True
    End If
    
    B1lb_BasyoN = ""
    If RTrim(iB1_BasyoCD) <> "" Then
       wSel = "SELECT * FROM 名称マスタ"
       wSel = wSel & " WHERE 区分名称 = '返却場所'"
       wSel = wSel & "   AND 値       = '" & Trim(iB1_BasyoCD.Text) & "'"

       If SYMRead(wSel) = True Then
          B1lb_BasyoN = SYM.値名称
       End If
    End If
'    Call iB1_Chk
'    Call iB1_BusyoCD.SetFocus
End Sub
'+-----------------------------+
'+  iB1_Chk
'+-----------------------------+
'Private Sub iB1_Chk()
'    With vsGrid
'        If ProcHB <> "B1" Then Exit Sub
''        If vsGrid.Row > RD_Count Then
''           If SyoriKB <> "DEL" Then iB1_BusyoCD.SetFocus
''           Exit Sub
''        End If
'        If Trim(iB1_BusyoCD) <> Trim(.TextMatrix(.Row, 5)) Or _
'           Trim(iB1_BasyoCD) <> Trim(.TextMatrix(.Row, 6)) Then
'           For gInt = 0 To 4: .Col = gInt: .CellBackColor = &HC0FFC0: Next  ' &HFFFFC0: Next
'        Else
'           .Col = 0: .CellBackColor = &H8000000F
'           For gInt = 1 To 4: .Col = gInt: .CellBackColor = .BackColor: Next
'        End If
'    End With
'End Sub
'+-----------------------------+
'+  B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
'    If Back_JTorcd.Tag = "S" Then
'        Call TextToVsGrid(Val(Input_Row.Tag))
'        Call Grid_RowChange(vbKeyUp)
'    End If
'    If Back_JTorcd.Tag = "E" Then
'        Call TextToVsGrid(Val(Input_Row.Tag))
'        Call Grid_RowChange(vbKeyDown)
'    End If
'    SendKeys "{Tab}"
End Sub
Private Sub B_TabDummy_LostFocus()
    On Error Resume Next
'    If iB1_STorcd.Enabled Then iB1_STorcd.SetFocus
    On Error GoTo 0
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
     Case "MNT"
          SyoriNM = "修正"
          K_Sykbnm.ForeColor = &HFF0000  '(青色)
     Case "DEL"
          SyoriNM = "削除"
          K_Sykbnm.ForeColor = &HFF&     '(赤色)
          W_Color = &HFF&
'''' Case "LOK"
''''      SyoriNM = "照会"
''''      K_Sykbnm.ForeColor = &HFF00FF   '(ピンク)
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
   PB_CAN.Visible = True        ' 2009.10.22 add : 処理名がある時は取消ボタン有効
   K_Sykbnm.Caption = SyoriNM
   K_Sykbnm.Visible = True
   Timer2.Enabled = True
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID
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

    H1Mode = False: H1Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1": H1Mode = True: H1Color = gPGotSel
       Case "B1": B1Mode = True: B1Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode: H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode: B1_Area1.BackColor = B1Color
    
    PB_Look.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_ADD.Visible = H1Mode         ' 2009.10.22 add
    PB_MNT.Visible = H1Mode         ' 2009.10.22 add
    PB_DEL.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    
    If SyoriKB <> "DEL" Then        ' 2009.10.22 add
       Call GridTextEnabled(B1Mode)
    End If                          ' 2009.10.22 add
    
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
        CisVsGrid.NumberSet = No    ' 2009.10.22 upd : All --> No
'       CisVsGrid.ReNumber          ' 2009.10.22 del
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        If SyoriKB = "DEL" Then
'           If Not DBDelete Then GoTo ReturnPress_Ed
'           Call HeadBodyClear("B")
'           ProcHB = "H1"
           ProcHB = "B1"
           Call DispChange(ProcHB)
           If Trim(vsGrid.TextMatrix(1, 1)) <> "" Then
              vsGrid.Row = 1
              Call vsGrid.Select(1, 1): DoEvents '              Call vsGrid_KeyDown(vbKeyDown)
           End If
'           SyoriKB = ""
           Call SyoriName(SyoriKB)
        Else
           ProcHB = "B1"
           Call DispChange(ProcHB)
           vsGrid.Row = 1
           Call VsGridToText(vsGrid.Row)
        End If
        GoTo ReturnPress_Ed
    End If
GoTo jmp0001
'
        If SyoriKB <> "DEL" Then
'            ProcHB = "B1"
'            Call DispChange(ProcHB)
'            vsGrid.Row = 1
'            Call GridTextEnabled(True)
            Call VsGridToText(vsGrid.Row)
'           If SyoriKB = "ADD" Then
'              gSL_Select = "SELECT MAX(値) AS 値 FROM 名称マスタ WHERE 区分名称 = '工程かんばん発行パターン'"
'              If SYMRead(gSL_Select) Then
'                 vsGrid.TextMatrix(1, 0) = Format(CisFun.Val2(SYM.値) + 1, "000")
'              Else
'                 vsGrid.TextMatrix(1, 0) = "001"
'              End If
'              '---------------------------------------' 2009.10.26 add start
'              With vsGrid
'                  .Rows = 2
'                  .Col = 5
'                  .Row = 1
'                  .CellBackColor = .BackColorFixed
'              End With
'              RD_Count = 1
'              Input_Row.Tag = 1
'              '---------------------------------------' 2009.10.26 add end
'           End If
'           Call VsGridToText(vsGrid.Row)
'           Call iB1_BusyoCD.SetFocus    ' 2009.10.22 add
'           Call iB1_BusyoCD_GotFocus    ' 2009.10.22 add
        End If
    GoTo ReturnPress_Ed
    
jmp0001:
    If ProcHB = "B1" Then
'---------------------------------------------------' 2009.10.23 upd start : 削除時の処理を追加
        If SyoriKB = "DEL" Then
           If Not DBDelete Then GoTo ReturnPress_Ed
        Else
           If Not Body1Chk Then GoTo ReturnPress_Ed
           If Not DBPut Then GoTo ReturnPress_Ed
        End If
'---------------------------------------------------' 2009.10.23 upd end
        Call HeadBodyClear("B")
        Call GridTextEnabled(False)                 ' 2009.10.22 add : 入力域非表示
        ProcHB = "H1"
        Call DispChange(ProcHB)
        SyoriKB = ""
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

    ' 部署
    If Trim(iH1_BusyoCD) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    部署を入力して下さい              "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_BusyoCD.SetFocus
       Exit Function
    End If
    Call iH1_BusyoCD_LostFocus
    If H1lb_BusyoN = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    部署　未登録！                    "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_BusyoCD.SetFocus
       Exit Function
    End If

    ' 返却場所
    If Trim(iH1_BasyoCD) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    返却場所を入力して下さい          "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_BasyoCD.SetFocus
       Exit Function
    End If
    Call iH1_BasyoCD_LostFocus
    If H1lb_BasyoN = "" Then
         With CisFun
             .MB_Lines = 4
             .MB_MSG(2) = "    返却場所　未登録！                "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
         End With
        iH1_BasyoCD.SetFocus
        Exit Function
     End If

    gSL_Select = "SELECT M1.値, M1.値名称, M1.英数字1"
    gSL_Select = gSL_Select & "     , M1.英数字2, T2.略称   AS 名称名"
    gSL_Select = gSL_Select & "     , M1.英数字3, M3.値名称 AS 検索名"
    gSL_Select = gSL_Select & "  FROM 名称マスタ M1"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ T2"
    gSL_Select = gSL_Select & "              ON  T2.取引先CD   = M1.英数字2"
    gSL_Select = gSL_Select & "              AND T2.取引先区分 = 1"             ' 取引先区分：発注
'   gSL_Select = gSL_Select & "              AND T2.社内区分   = 1"             ' 社内区分　：社内  ' 2009.10.26 del
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ   M3"
    gSL_Select = gSL_Select & "              ON  M3.区分名称   = '返却場所'"
    gSL_Select = gSL_Select & "              AND M3.値         = M1.英数字3"
    gSL_Select = gSL_Select & " WHERE M1.区分名称 = '工程かんばん発行パターン'"
    gSL_Select = gSL_Select & "   AND ISNULL(M1.値,'') <> ''"
    gSL_Select = gSL_Select & "   AND M1.値名称   = '" & Trim(iH1_BusyoCD) & "'"
    gSL_Select = gSL_Select & "   AND M1.英数字1  = '" & Trim(iH1_BasyoCD) & "'"
    gSL_Select = gSL_Select & " ORDER BY M1.値"

    If DBInput("INV") Then
        If SyoriKB = "" Then
            SyoriKB = "MNT"
            Call SyoriName(SyoriKB)
        End If
    Else
        If SyoriKB = "DEL" Or SyoriKB = "MNT" Then      ' 2009.10.22 upd : "MNT" 追加
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    該当データが存在しません！        "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
            iH1_BusyoCD.SetFocus
            Exit Function
        End If
        SyoriKB = "ADD"
        Call SyoriName(SyoriKB)
    End If
       
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Dim wSel    As String
    
    Body1Chk = False

    If Trim(iB1_BusyoCD) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    部署１を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_BusyoCD.SetFocus
        Exit Function
    End If
    Call iB1_BusyoCD_LostFocus
    If Trim(B1lb_BusyoN) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    部署１　未入力                    "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_BusyoCD.SetFocus
        Exit Function
    End If
    
    If Trim(iB1_BasyoCD) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    返却場所１を入力して下さい        "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_BasyoCD.SetFocus
        Exit Function
    End If
    Call iB1_BasyoCD_LostFocus
    If Trim(B1lb_BasyoN) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    返却場所１　未入力                "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_BasyoCD.SetFocus
        Exit Function
    End If

    With vsGrid
        .TextMatrix(.Row, 1) = Trim(iB1_BusyoCD)
        .TextMatrix(.Row, 2) = Trim(B1lb_BusyoN)
        .TextMatrix(.Row, 3) = Trim(iB1_BasyoCD)
        .TextMatrix(.Row, 4) = Trim(B1lb_BasyoN)
        If (.TextMatrix(.Row, 1) = Trim(.TextMatrix(.Row, 5))) And _
           (.TextMatrix(.Row, 3) = Trim(.TextMatrix(.Row, 6))) Then
            .TextMatrix(.Row, 7) = ""
            .Col = 0: .CellBackColor = &H8000000F
            For gInt = 1 To 4: .Col = gInt: .CellBackColor = .BackColor: Next
        Else
            .TextMatrix(.Row, 7) = "X"
            For gInt = 0 To 4: .Col = gInt: .CellBackColor = &HC0FFC0: Next
        End If
    End With

'ﾒｯｾｰｼﾞﾎﾞｯｸｽ設定
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
'<< 行を詰める >>
    With vsGrid
'-*-*-*-*-*-*-*-*- 2006/09/05 Del Start -*-*-*-*-*-*-*-*-*-*-*-*-
        For gCnt1 = 1 To .Rows - 1
            If Not BodyInput(gCnt1) Then
                For gCnt2 = gCnt1 + 1 To .Rows - 1
                    If BodyInput(gCnt2) Then
                        For gCnt3 = 1 To .Cols - 1
                            .TextMatrix(gCnt1, gCnt3) = .TextMatrix(gCnt2, gCnt3)
                            .TextMatrix(gCnt2, gCnt3) = ""
                        Next gCnt3
                        Exit For
                    End If
                Next gCnt2
            End If
        Next gCnt1
'-*-*-*-*-*-*-*-*- 2006/09/05 Del End   -*-*-*-*-*-*-*-*-*-*-*-*-
'
        For gCnt1 = 1 To .Rows - 1
            If Not BodyInput(gCnt1) Then Exit For
        '<< 部署１入力チェック >>
            If RTrim(.TextMatrix(gCnt1, 1)) = "" Then
                CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　部署１　未入力！】                "
                CisFun.MBOX
                Set ErrObj = iB1_BusyoCD
                GoTo Body1Chk_Err
            End If
            
''''            .TextMatrix(gCnt1, 2) = ""
''''            If TorNmGet(.TextMatrix(gCnt1, 1), 9) Then
''''               B1lb_BusyoN = RTrim(TRM.略称)  ' TRM_RName
''''            End If
            If RTrim(.TextMatrix(gCnt1, 2)) = "" Then
                CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　部署１　未登録！】                "
                CisFun.MBOX
                Set ErrObj = iB1_BusyoCD
                GoTo Body1Chk_Err
            End If

        '<< 返却場所１チェック >>
            If RTrim(.TextMatrix(gCnt1, 3)) = "" Then
                CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　返却場所１　未入力！】            "
                CisFun.MBOX
                Set ErrObj = iB1_BusyoCD
                GoTo Body1Chk_Err
            End If
            
''''            .TextMatrix(gCnt1, 4) = ""
''''            wSel = "SELECT * FROM 名称マスタ"
''''            wSel = wSel & " WHERE 区分名称 = '返却場所'"
''''            wSel = wSel & "   AND 値       = '" & Trim(iB1_BasyoCD.Text) & "'"
''''            If SYMRead(wSel) = True Then
''''               B1lb_BasyoN = SYM.値名称
''''            End If
            If RTrim(.TextMatrix(gCnt1, 2)) = "" Then
                CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　返却場所１　未登録！】            "
                CisFun.MBOX
                Set ErrObj = iB1_BusyoCD
                GoTo Body1Chk_Err
            End If
        
''''        '<< 返却場所１チェック >>
''''            If Not TorNmGet(.TextMatrix(gCnt1, 1), 8) Then
''''                CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　取引先マスタ　未登録！】            "
''''                CisFun.MBOX
''''                Set ErrObj = iB1_JTorcd
''''                GoTo Body1Chk_Err
''''            End If
''''        '<< 取引先受入マスタ存在チェック >>
''''            TUM.取引先CD = RTrim$(.TextMatrix(gCnt1, 1))
''''            TUM.取引先区分 = 1
''''            TUM.受入 = RTrim$(.TextMatrix(gCnt1, 3))
''''            If Not TUMRead("", 1) Then
''''                CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　取引先受入マスタ　未登録！】            "
''''                CisFun.MBOX
''''                Set ErrObj = iB1_JUke
''''                GoTo Body1Chk_Err
''''            End If
''''        '<< 契約取引先入力必須 >>
''''            If RTrim(.TextMatrix(gCnt1, 4)) = "" Then
''''                CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　契約取引先　未入力！】            "
''''                CisFun.MBOX
''''                Set ErrObj = iB1_KTorcd
''''                GoTo Body1Chk_Err
''''            End If
        '<< 重複チェック >>
            For gCnt2 = gCnt1 + 1 To .Rows - 1
''''                If RTrim(.TextMatrix(gCnt1, 1)) = "" Then Exit For
''''                If RTrim(.TextMatrix(gCnt1, 1)) = RTrim(.TextMatrix(gCnt2, 1)) And _
''''                   RTrim(.TextMatrix(gCnt1, 3)) = RTrim(.TextMatrix(gCnt2, 3)) And _
''''                   RTrim(.TextMatrix(gCnt1, 9)) = "" And RTrim(.TextMatrix(gCnt2, 9)) = "" Then
''''                    CisFun.MB_MSG(1) = "    　自社取引先が重複しています。            "
''''                    CisFun.MB_MSG(3) = "    " & gCnt1 & "行目と " & gCnt2 & "行目            "
''''                    CisFun.MBOX
''''                    Set ErrObj = iB1_JTorcd
''''                    gCnt1 = gCnt2
''''                    GoTo Body1Chk_Err
''''                End If
            
                If RTrim(.TextMatrix(gCnt1, 1)) = RTrim(.TextMatrix(gCnt2, 1)) And _
                   RTrim(.TextMatrix(gCnt1, 3)) = RTrim(.TextMatrix(gCnt2, 3)) Then     ''''And _
''''                   RTrim(.TextMatrix(gCnt1, 4)) = RTrim(.TextMatrix(gCnt2, 4)) And _
''''                   RTrim(.TextMatrix(gCnt1, 5)) = RTrim(.TextMatrix(gCnt2, 5)) And _
''''                   RTrim(.TextMatrix(gCnt1, 6)) = RTrim(.TextMatrix(gCnt2, 6)) And _
''''                   RTrim(.TextMatrix(gCnt1, 7)) = RTrim(.TextMatrix(gCnt2, 7)) And _
''''                   RTrim(.TextMatrix(gCnt1, 8)) = RTrim(.TextMatrix(gCnt2, 8)) Then
                    CisFun.MB_MSG(1) = "    　内容が重複しています。            "
                    CisFun.MB_MSG(3) = "    " & gCnt1 & "行目と " & gCnt2 & "行目            "
                    CisFun.MBOX
                    Set ErrObj = iB1_BusyoCD
                    gCnt1 = gCnt2
                    GoTo Body1Chk_Err
                End If
            Next gCnt2
            
        Next gCnt1
    End With
    
    Body1Chk = True
    Exit Function
Body1Chk_Err:
    vsGrid.Row = gCnt1
    Call VsGridToText(vsGrid.Row)
    Call GridTextEnabled(True)
    ErrObj.SetFocus
End Function
'+----------------------+
'+      入 力 判 定
'+----------------------+
Private Function BodyInput(RowNo As Long)
    BodyInput = False

    With vsGrid
        If RTrim(.TextMatrix(RowNo, 0)) = "" And _
           RTrim(.TextMatrix(RowNo, 1)) = "" And _
           RTrim(.TextMatrix(RowNo, 3)) = "" Then Exit Function
    End With

    BodyInput = True
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
        
    If Dsp_Flg = 0 Then Call HeadBodyClear("B")
    
    RD_Count = 0
    If Not SYMRead(gSL_Select, 0, 1) Then
        Call SYMClose(1)
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
        RD_Count = CisDB.RecordCount
        Call BodySet
      ''vsGrid.Rows = RD_Count + 1
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

    With SYM
''''        iH1_Torcd = RTrim(.契約先)
''''        Call iH1_Torcd_LostFocus
''''        iH1_DaikoK = RTrim(.代行契約先)                             ' 2007/12/17 ADD
''''        Call iH1_DaikoK_LostFocus                                   ' 2007/12/17 ADD
''''
''''        iB1_STorcd = RTrim(.支給元)
''''        iB1_TyokTorcd = RTrim(.支給元2)
''''
''''        If Trim(B1lb_HiJisya) = "" Then
''''           B1lb_HiJisya = mJisya
''''        End If
''''
'''''
''''        mSTorcd = RTrim(.支給元)
''''        mTyokTorcd = RTrim(.支給元2)
        
    End With
'Grid表示
    Call GridSet
    Call SYMClose(1)
        
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
    Dim wSel    As String
    Dim wCnt    As Integer
    
    DBDelete = False

'  ( 選択確認)
    wSel = ""
    With vsGrid
        For wCnt = 1 To .Rows - 1
            If Trim(.TextMatrix(wCnt, 0)) = "" Then Exit For
            If Trim(.TextMatrix(wCnt, 7)) <> "" Then
               If wSel = "" Then
                  wSel = "DELETE FROM 名称マスタ"
                  wSel = wSel & " WHERE 区分名称 = '工程かんばん発行パターン' AND "
                  wSel = wSel & "     ( 値 = '" & Trim(.TextMatrix(wCnt, 0)) & "'"
               Else
                  wSel = wSel & "    OR 値 = '" & Trim(.TextMatrix(wCnt, 0)) & "'"
               End If
            End If
        Next
    End With
    If wSel <> "" Then wSel = wSel & "     )"
    
    If wSel = "" Then
       With CisFun
           .MB_Lines = 4
           .MB_MSG(2) = "     削除する登録№を選択してください     "
           .MB_Title = "削除処理"
           .MB_Button = OK
           .MBOX
        End With
        Exit Function
    End If

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
''''        ' 取引先変換マスタ
''''        .SQL = "DELETE FROM 取引先変換マスタ "
''''        .SQL = .SQL & " WHERE 契約先 = '" & RTrim(iH1_Torcd) & "'"
''''        .SQL = .SQL & "   AND 代行契約先 = '" & RTrim(iH1_DaikoK) & "'"             '2007/12/17 ADD
        .SQL = wSel                                                         ''' 2009.10.23 add
        .DBExec
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

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
Private Function DBPut(Optional pKB As Byte = 0)
    DBPut = False
    Dim AddFlg  As Boolean
    Dim wYDJun As Byte
    Dim wCnt    As Integer
    Dim wSel    As String
    Dim wTNo    As Integer

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    On Error GoTo DBPut_Err
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

    
''''''''With vsGrid                                         '''' TEST 2009.10.23
''''''''    .TextMatrix(.Row, 1) = Trim(iB1_BusyoCD)
''''''''    .TextMatrix(.Row, 2) = Trim(B1lb_BusyoN)
''''''''    .TextMatrix(.Row, 3) = Trim(iB1_BasyoCD)
''''''''    .TextMatrix(.Row, 4) = Trim(B1lb_BasyoN)
''''''''    .TextMatrix(.Row, 7) = "X"
''''''''End With
    
    If SyoriKB = "MNT" Then
        wSel = ""
        With vsGrid
            For wCnt = 1 To .Rows - 1
                If Trim(.TextMatrix(wCnt, 0)) = "" Then Exit For
'                If Trim(.TextMatrix(wCnt, 7)) <> "" Then
                   If wSel = "" Then
                      wSel = "DELETE FROM 名称マスタ"
                      wSel = wSel & " WHERE 区分名称 = '工程かんばん発行パターン' AND "
                      wSel = wSel & "     ( 値 = '" & Trim(.TextMatrix(wCnt, 0)) & "'"
                   Else
                      wSel = wSel & "    OR 値 = '" & Trim(.TextMatrix(wCnt, 0)) & "'"
                   End If
'                End If
            Next
        End With
        If wSel <> "" Then wSel = wSel & "     )"
        
        If wSel = "" Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "     修正対象データがありません           "
               .MB_Title = "修正処理"
               .MB_Button = OK
               .MBOX
            End With
            Exit Function
        End If

        If pKB = 0 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     更新処理を行います      "
                .MB_Title = "更新処理"
                .MB_MSG(3) = "     よろしいですか？      "
                .MB_Button = OK_CAN
                If Not .MBOX Then
                    Call VsGridToText(Val(Input_Row.Tag))
                    Call GridTextEnabled(True)
                    iB1_BusyoCD.SetFocus
                    Exit Function
                End If
            End With
        End If

        With CisDB
            On Error GoTo DBPut_Err
    ' ( 修正データ削除処理 )
            ' 名称マスタ
            .SQL = wSel
            .DBExec
            On Error GoTo 0
        End With
        '---　ここまで SyoriKB = "MNT" のとき　--------
    '
    Else    ' SyoriKB = "ADD" のときはこちら
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     登録処理を行います      "
            .MB_Title = "新規登録処理"
            .MB_MSG(3) = "     よろしいですか？      "
            .MB_Button = OK_CAN
            If Not .MBOX Then
                Call VsGridToText(Val(Input_Row.Tag))
                Call GridTextEnabled(True)
                iB1_BusyoCD.SetFocus
                Exit Function
            End If
        End With
    End If
    
    ' 最終登録№を取得する
    gSL_Select = "SELECT MAX(値) AS 値 FROM 名称マスタ "
    gSL_Select = gSL_Select & " WHERE 区分名称 = '工程かんばん発行パターン'"
    If SYMRead(gSL_Select) Then
       wTNo = CisFun.Val2(SYM.値)
    Else
       wTNo = 0
    End If
    
    ' 追加処理(修正も追加で実施)
    With vsGrid
        For wCnt = 1 To .Rows - 1
             If Trim(.TextMatrix(wCnt, 1)) = "" Then Exit For
'            If Trim(.TextMatrix(wCnt, 7)) <> "" Then
               Call ItemsClearSYM
               With SYM
                    .区分名称 = "工程かんばん発行パターン"
                   '.区分タイプ = ""
                    If Trim(vsGrid.TextMatrix(wCnt, 0)) <> "" Then
                       .値 = vsGrid.TextMatrix(wCnt, 0)
                    Else
                       wTNo = wTNo + 1
                       .値 = Format(wTNo, "000")
                    End If
                    .値名称 = RTrim(iH1_BusyoCD)
                    .区分桁数 = 3
                    .英数字1 = RTrim(iH1_BasyoCD)
                    .英数字2 = RTrim(vsGrid.TextMatrix(wCnt, 1))    ' RTrim(iB1_BusyoCD)
                    .英数字3 = RTrim(vsGrid.TextMatrix(wCnt, 3))    ' RTrim(iB1_BasyoCD)
                    .制御区分 = "X"
               End With
               Call SYMInsert       ' 名称マスタ書込
'            End If
        Next
    End With

    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'*------------------------------------------------------------------------------*
'*      退避データと入力テキストとのチェック ( 異なっていれば更新をうながす )       *
'*------------------------------------------------------------------------------*
Private Function SaveDateCheck()
'-----------------------------------------------------------------------------------' 2009.10.22 del start
''''    SaveDateCheck = False
''''
''''    Static swTopR           As Integer
''''    Static swRow            As Integer
''''
''''    swTopR = vsGrid.TopRow
''''    swRow = vsGrid.Row
''''
''''    If Not Body1Chk Then Exit Function
''''    Call DBPut(1)
''''    Input_Row.Visible = False
''''    Call HeadBodyClear("B")
''''    Call Head1Chk
''''    Input_Row.Visible = True
''''    vsGrid.TopRow = swTopR
''''
''''    If vsGrid.Rows <= swRow Then swRow = vsGrid.Rows - 1
''''    vsGrid.Row = swRow
''''    Call VsGridToText(vsGrid.Row)
''''    DoEvents
''''    iB1_JTorcd.SetFocus
''''
''''
'''''    With vsGrid
'''''        If mTorcd(tRow) = RTrim$(.TextMatrix(tRow, 1)) And _
'''''            mUkeir(tRow) = RTrim$(.TextMatrix(tRow, 3)) And _
'''''            mKTorcd(tRow) = RTrim$(.TextMatrix(tRow, 10)) And _
'''''            mKTorkj(tRow) = RTrim$(.TextMatrix(tRow, 11)) And _
'''''            mSTorcd2(tRow) = RTrim$(.TextMatrix(tRow, 12)) And _
'''''            mSTorkj2(tRow) = RTrim$(.TextMatrix(tRow, 13)) And _
'''''            mSTorUke2(tRow) = RTrim$(.TextMatrix(tRow, 14)) And _
'''''            mNTorcd(tRow) = RTrim$(.TextMatrix(tRow, 15)) And _
'''''            mNTorkj(tRow) = RTrim$(.TextMatrix(tRow, 16)) And _
'''''            mNUkeir(tRow) = RTrim$(.TextMatrix(tRow, 7)) And _
'''''            mNTorNm(tRow) = RTrim$(.TextMatrix(tRow, 8)) Then
'''''                GoTo SaveDateCheck_ED
'''''        End If
'''''    End With
'''''
'''''    With CisFun
'''''        .MB_Lines = 5
'''''        .MB_MSG(1) = "      【" & StrConv(tRow, vbWide) & "行目の内容が変更されています。】        "
'''''        .MB_MSG(2) = "     品番選択の前に更新を行う必要があります。         "
'''''        .MB_MSG(4) = "              よろしいですか？      "
'''''        .MB_Title = "更新処理"
'''''        .MB_Button = OK_CAN
'''''        If Not .MBOX Then
'''''            iB1_JTorcd.SetFocus
'''''            Exit Function
'''''        End If
'''''    End With
''''
'-----------------------------------------------------------------------------------' 2009.10.22 del end
SaveDateCheck_ED:
    SaveDateCheck = True
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
        Call .FixedSet(中中, 中中, あり, 6, 0, "登録№")
        Call .FixedSet(中中, 左中, あり, 7, 1, "部署")
        Call .FixedSet(中中, 左中, あり, 30, 2, "１")
        Call .FixedSet(中中, 左中, あり, 3, 3, "返")
        Call .FixedSet(中中, 左中, あり, 30, 4, "１")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 中中, あり, 0, 5, "x契約取引先")
        Call .FixedSet(中中, 中中, あり, 0, 6, "x契約取引工場")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 7, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
       Call .PatanSet(0, "0,1,2,3,4")          ' 2009.10.22 del : これを実行すると何故かカラムの順序が入れ替わるため
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
    
    Ttl_Kensu = RD_Count

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
    
    ReDim mBusyoCD(Ttl_Kensu + 1)
    ReDim mBasyoCD(Ttl_Kensu + 1)
    
    Do Until Not SYM_RDSTS
        'Grid Row,Col
        ID = ID + 1
        With vsGrid
            .TextMatrix(ID, 0) = Mid(SYM.値, 1, 3)     ' 念のため先頭から３桁のみ抽出      ' ID
            .TextMatrix(ID, 1) = SYM.英数字2 + Space(7 - CisFun.Len2(SYM.英数字2))
            .TextMatrix(ID, 2) = RTrim(SYM.名称名)
            .TextMatrix(ID, 3) = SYM.英数字3 + Space(3 - CisFun.Len2(SYM.英数字3))
            .TextMatrix(ID, 4) = RTrim(SYM.検索名)
            .TextMatrix(ID, 5) = .TextMatrix(ID, 1)
            .TextMatrix(ID, 6) = .TextMatrix(ID, 3)
            .TextMatrix(ID, 7) = ""    ' Select
        End With
        '初期情報退避
                
        mBusyoCD(ID) = RTrim(SYM.英数字2)
        mBasyoCD(ID) = RTrim(SYM.英数字3)
        Call SYMReadNext(1)
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
    Input_Row.Visible = EnabledFlg
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
        If gInt + 1 > vsGrid.Rows - 1 Then
            If RTrim$(vsGrid.TextMatrix(gInt, 1)) <> "" Then
                vsGrid.Rows = vsGrid.Rows + 1
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
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    Dim wSel    As String
    
    With vsGrid
'       >> 部署１
        .TextMatrix(tRow, 1) = iB1_BusyoCD + Space(7 - CisFun.Len2(iB1_BusyoCD))
        .TextMatrix(tRow, 2) = ""
        If RTrim(iB1_BusyoCD) <> "" Then
            If TorNmGet(iB1_BusyoCD, 9) Then
                .TextMatrix(tRow, 2) = RTrim(TRM.略称)
            End If
        End If
'       >> 返却場所１
        .TextMatrix(tRow, 3) = iB1_BasyoCD + Space(3 - CisFun.Len2(iB1_BasyoCD))
        .TextMatrix(tRow, 4) = B1lb_BasyoN
        If RTrim(iB1_BasyoCD) <> "" Then
           wSel = "SELECT * FROM 名称マスタ"
           wSel = wSel & " WHERE 区分名称 = '返却場所'"
           wSel = wSel & "   AND 値       = '" & Trim(iB1_BasyoCD.Text) & "'"
    
           If SYMRead(wSel) = True Then
              B1lb_BasyoN = SYM.値名称
           End If
        End If
    End With
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With vsGrid
        Input_Row.Tag = tRow
        Input_Row.Move Input_Row.Left, vsGrid.CellTop, Input_Row.Width, vsGrid.CellHeight
        
        ' 部署コード１
        iB1_BusyoCD = RTrim(.TextMatrix(tRow, 1))
        ' 部署名１
        B1lb_BusyoN = RTrim(.TextMatrix(tRow, 2))
        ' 返却場所コード１
        iB1_BasyoCD = RTrim(.TextMatrix(tRow, 3))
        '> 返却場所名２
        B1lb_BasyoN = RTrim(.TextMatrix(tRow, 4))
    End With
End Sub
Private Sub vsGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If SyoriKB = "DEL" Then
       If ProcHB = "B1" And SyoriKB = "DEL" And KeyCode = vbKeySpace Then
          Call CisVsGrid.vsColor
       End If
       Exit Sub
    End If
    
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
    
    Call TextToVsGrid(Val(Input_Row.Tag))
    Call VsGridToText(vsGrid.Row)
End Sub
Private Sub VSGrid_Click()
    Dim wCnt    As Integer
    
    If SyoriKB = "DEL" Then
       If Trim(vsGrid.TextMatrix(vsGrid.Row, 1)) = "" Then Exit Sub
'       With vsGrid
''.CellAlignment
'           If Trim(.TextMatrix(.Row, 0)) = "" Then Exit Sub
'           If Trim(.TextMatrix(.Row, 7)) = "" Then
'              .TextMatrix(.Row, 7) = "X"
'          ''''Call CisVsGrid.vsColor(個別, &HFFFFC0)
'              For wCnt = 0 To 4: .Col = wCnt: .CellBackColor = &HC0FFC0: Next
'
'           Else
'              .TextMatrix(.Row, 7) = ""
'          ''''Call CisVsGrid.vsColor(個別, &HC0FFC0)
'              .Col = 0: .CellBackColor = &H8000000F
'              For wCnt = 1 To 4: .Col = wCnt: .CellBackColor = .BackColor: Next
'           End If
'       End With
       Call CisVsGrid.vsColor
       Kensu1 = Format(CisVsGrid.SelectCount, "    0")
       Exit Sub
    End If

    Call TextToVsGrid(Val(Input_Row.Tag))
'    If RTrim$(vsGrid.TextMatrix(vsGrid.Row, 0)) = "" Then Exit Sub
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

'+---------------------------------------+
'+      名称マスタ初期設定
'+---------------------------------------+
Private Sub MeisyoInit()
'   >> 後工程符号管理
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "後工程符号管理"
        .区分桁数 = 5
        .値 = "1097"
        .値名称 = "アラコ"
        .英数字1 = "AR"
        .作成者 = gTanto
        If Not SYMInitCreate(True) Then
            .値 = "1613"
            .値名称 = "セントラル"
            .英数字1 = "CN"
            Call SYMInitCreate
            .値 = "1875"
            .値名称 = "ダイハツ"
            .英数字1 = "DH"
            Call SYMInitCreate
            .値 = "2474"
            .値名称 = "岐阜車"
            .英数字1 = "GI"
            Call SYMInitCreate
            .値 = "2625"
            .値名称 = "日野"
            .英数字1 = "HN"
            Call SYMInitCreate
            .値 = "3230"
            .値名称 = "関自"
            .英数字1 = "KN"
            Call SYMInitCreate
            .値 = "7371"
            .値名称 = "織機"
            .英数字1 = "KY"
            Call SYMInitCreate
            .値 = "7388"
            .値名称 = "トヨ車"
            .英数字1 = "TY"
            Call SYMInitCreate
            .値 = "8482"
            .値名称 = "トヨタ九州"
            .英数字1 = "TMK"
            Call SYMInitCreate
            .値 = "7364"
            .値名称 = "トヨタ紡織"
            .英数字1 = "TB"
            Call SYMInitCreate
        End If
    End With
End Sub

