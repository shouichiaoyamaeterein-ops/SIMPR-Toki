VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Begin VB.Form CKH0410 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "内示かんばん情報照会"
   ClientHeight    =   11145
   ClientLeft      =   690
   ClientTop       =   2475
   ClientWidth     =   15930
   ControlBox      =   0   'False
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
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11145
   ScaleWidth      =   15930
   WindowState     =   2  '最大化
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4935
      ScaleHeight     =   615
      ScaleWidth      =   4530
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   675
      Width           =   4590
      Begin VB.Frame Frm_Sort 
         BackColor       =   &H00808000&
         Caption         =   "《分類》"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   60
         TabIndex        =   14
         Top             =   30
         Width           =   1635
         Begin VB.OptionButton Op_Sort 
            BackColor       =   &H00808000&
            Caption         =   "降順"
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   1
            Left            =   870
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   180
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.OptionButton Op_Sort 
            BackColor       =   &H00808000&
            Caption         =   "昇順"
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   180
            Width           =   675
         End
      End
      Begin Cis3D_v60.CIS3D PNL_Kensu 
         Height          =   465
         Left            =   1740
         Top             =   90
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   820
         ForeColor       =   192
         Caption         =   "件数"
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
         cBoderWidth     =   3
         cFont3D         =   1
         cAlingnment     =   1
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D Lb_Kensu 
            Height          =   345
            Left            =   675
            Top             =   60
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   609
            BackColor       =   16777152
            Caption         =   "99,999/99,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
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
      Left            =   9585
      ScaleHeight     =   660
      ScaleWidth      =   5115
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   615
      Width           =   5175
      Begin CisBtn_60.CisBtn PB_CAN 
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
         Left            =   4245
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
      Begin CisBtn_60.CisBtn PB_DSP 
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
         BtnType         =   12
         BtnCaption      =   "照会"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   16711935
         BtnShadowColor  =   16777215
         PFCaption       =   "( F4 )"
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
         PFBackColor     =   16744703
      End
      Begin CisBtn_60.CisBtn PB_MNT 
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
      Height          =   1005
      Left            =   345
      ScaleHeight     =   945
      ScaleWidth      =   14355
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1425
      Width           =   14415
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   855
         Left            =   60
         Top             =   60
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   1508
         ForeColor       =   16711680
         Caption         =   "内示年月"
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
         Begin CisYM_V60.CisYM iH1_NJYm 
            Height          =   360
            Left            =   240
            TabIndex        =   7
            Top             =   345
            Width           =   1380
            _ExtentX        =   2328
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
            cSize           =   -1  'True
            cChkResult      =   0   'False
         End
      End
      Begin Cis3D_v60.CIS3D BLB_HINBN 
         Height          =   855
         Left            =   6375
         Top             =   60
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1508
         ForeColor       =   16711680
         Caption         =   "品  番"
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
            Left            =   120
            TabIndex        =   2
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
         Height          =   855
         Index           =   1
         Left            =   1950
         Top             =   60
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   1508
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
            Left            =   1080
            Top             =   330
            Width           =   2460
            _ExtentX        =   4339
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
         Height          =   855
         Left            =   5610
         Top             =   60
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1508
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
            Left            =   165
            TabIndex        =   1
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
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   855
         Left            =   12435
         Top             =   60
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   1508
         ForeColor       =   16711680
         Caption         =   "増減枚数差"
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
         Begin Cis3D_v60.CIS3D CIS3D16 
            Height          =   345
            Left            =   240
            Top             =   345
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   609
            Caption         =   "±"
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
         End
         Begin CisText_V60.CisText iH1_Maisu 
            Height          =   360
            Left            =   465
            TabIndex        =   6
            Top             =   330
            Width           =   390
            _ExtentX        =   688
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
         Begin Cis3D_v60.CIS3D CIS3D17 
            Height          =   345
            Left            =   900
            Top             =   345
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   609
            Caption         =   "枚以上"
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
      End
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   855
         Left            =   9750
         Top             =   60
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   1508
         ForeColor       =   16711680
         Caption         =   "増減区分"
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
         Begin VB.OptionButton OpH1_ZouGen 
            BackColor       =   &H00C0C0C0&
            Caption         =   "減"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   1845
            TabIndex        =   5
            Top             =   375
            Width           =   600
         End
         Begin VB.OptionButton OpH1_ZouGen 
            BackColor       =   &H00C0FFC0&
            Caption         =   "全て"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   90
            TabIndex        =   3
            Top             =   375
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OpH1_ZouGen 
            BackColor       =   &H00C0C0C0&
            Caption         =   "増"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   1035
            TabIndex        =   4
            Top             =   375
            Width           =   600
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   8385
      Left            =   330
      ScaleHeight     =   8325
      ScaleWidth      =   14370
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2490
      Width           =   14430
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   8265
         Left            =   45
         Top             =   45
         Width           =   14280
         _ExtentX        =   25188
         _ExtentY        =   14579
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
         Begin TrueOleDBGrid60.TDBGrid DBGrid1 
            Bindings        =   "CKH0410.frx":0000
            Height          =   8115
            Left            =   45
            OleObjectBlob   =   "CKH0410.frx":0015
            TabIndex        =   12
            Top             =   60
            Width           =   14130
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
      TabIndex        =   8
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
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 内示かんばん情報照会  】"
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
      cPositionX      =   1600
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   16711935
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   635
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "CKH0410"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   内示かんばん情報照会
'**       フォームID    :   CKH0410
'**       処理概要      :
'**
'**       作  成  日    :   2004/04/09  By CIS
'**       変  更  日    :   2006/03/16  By CIS かんばん管理入力(増減･平準化)呼出変更
'**                                                 ( CKH0110 → CKH0170 )
'**                                            照会BODY入力時、選択ﾎﾞﾀﾝを非表示に修正
'**       変  更  日    :   2014/02/04  By CIS WIN7対応
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(照会･修正)
    Dim SyoriNM         As String       ' 処理名称(照会･修正)
    
    Dim MaxCnt          As Long
    Dim NowCnt          As Long
    Dim mSL_SelectW     As String       ' SQL SELECT(Where)条件文
    Dim mSL_OrderA      As String       ' SQL SELECT(Where)条件文
    Dim mSL_OrderD      As String       ' SQL SELECT(Where)条件文
    
    Dim CisFrm          As CISFormContorl
    
    Dim mGridMoveCol    As Boolean
    
    Dim mOPGotColor     As Long
    Dim mOPLostColor    As Long
    '--------------------------------------- 2006/03/16 Insert
    Dim mBaseStartDay   As Integer
    Dim mStartDay       As Integer         '開始日
    Dim mStartYMD       As String
    Dim mEndYMD         As String
    '--------------------------------------- 2006/03/16 Insert End
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    If RTrim(Rv_Call3) = "" Then
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
'
    Dummy.Left = -1000
    Dummy.Enabled = False
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 時刻表示間隔
    Timer2.Enabled = True    ' 時刻表示開始

    mOPGotColor = OpH1_ZouGen(0).BackColor
    mOPLostColor = OpH1_ZouGen(1).BackColor

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = "DSP"
    Call SyoriName(SyoriKB)
    
    OpH1_ZouGen(0).Value = True
    Op_Sort(0).Value = True
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    If RTrim(Rv_Call3) = "" Then Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
' ADOｺﾝﾄﾛｰﾙ　接続情報　セット
    Adodc1.ConnectionString = CisDB.ConnectString
    '-----2014/02/04 EDIT START-----
    gStr = CisDB.ConnectString
    gConnect = "server=" & CisDB.Server & ";database=" & CisDB.DBName & ";User ID=" & CisDB.User & ";password=" & CisDB.PassWord(GetSetting(gRegKey, "Environment", "DrowsSap", ""))
    
    gStr = Replace(gStr, "Use Procedure", "Extended Properties=" & gConnect & ";Use Procedure")
    
    Adodc1.ConnectionString = gStr
    '-----2014/02/04 EDIT END-------
    
'------------------------------------------------------ 2006/03/16 Insert
'名称マスタ【基本開始日】獲得
    If MeisyoGet("かんばん増減基本開始日", "1") Then
       mBaseStartDay = CisFun.Val2(SYM.英数字1)
    Else
       mBaseStartDay = 1
    End If
'------------------------------------------------------ 2006/03/16 Insert End

' 内示展開処理年月
    Call GetNJProc_YM(3)
    iH1_NJYm = Trim(gn_HNYM)
    
    FormAct = Not FormAct
'
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》                 +
'+-------------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    If RTrim(Rv_Call3) = "" Then Cancel = 1
    Call PB_END_Click
End Sub
'+-------------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》         +
'+-------------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    Select Case Key_Code
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click  '【修正】
           Case vbKeyF4:       If PB_DSP.Visible Then Call PB_DSP_Click  '【照会】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_Can_Click    '【取消】
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

    If RTrim(Rv_Call3) = "" Then
    '+---------------------+
    '+   ﾃﾞｰﾀﾍﾞｰｽ切断
    '+---------------------+
        Call CisDB.DBDISConnect
    '+---------------------+
    '+   ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
    '+---------------------+
        Call CisFrm.UnLoadDisp
        End
    Else
        Me.Hide
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
Private Sub PB_Can_Click()
   Call CisFrm.MousePT(11)
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call DBGridClr
      Call SyoriName(SyoriKB)
      iH1_Torcd.SetFocus
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
   Call CisFrm.MousePT(1)
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
'+    照　会  キ ー(F4)       +
'+----------------------------+
Private Sub PB_DSP_Click()
    If SyoriKB <> "DSP" Then
        SyoriKB = "DSP"
    Else
        SyoriKB = ""
    End If
    Call SyoriName(SyoriKB)
    iH1_Torcd.SetFocus
End Sub
'+----------------------------+
'+    検　索  キ ー(F7)        +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_UTorcd: Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
End Sub
'+-----------------------------+
'+       取引先検索             +
'+-----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 8
    
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
    RV_Call = "C"
    RV_Torcd = iH1_Torcd
    RV_TorKb = 8
    RV_Kojyo = ""
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
        iH1_Ukeir = RV_Ukeir
    End If
End Sub
'+----------------------------+
'+      品番検索              +
'+----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
    RVI_Torcd = RTrim(iH1_Torcd)
    RVI_Ukeir = RTrim(iH1_Ukeir)
    RVI_SK = 0
    RVI_Hinmoku = 1

    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iH1_Hinbn = RV_Hinbn
        iH1_Torcd = RV_TorcdK
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RV_Ukeir
    End If
End Sub
'+------------------------------------+
'+  iH1_Torcd(Got_Focus)
'+------------------------------------+
Private Sub iH1_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
'+------------------------------------+
'+  iH1_Torcd(Lost_Focus)
'+------------------------------------+
Private Sub iH1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Visible = False
        PB_Look.Tag = ""
    End If
    H1lb_Tornm = ""
    If TorNmGet(iH1_Torcd, 8) Then
        H1lb_Tornm = RTrim(TRM.略称)
    End If
End Sub
'+------------------------------------+
'+  iH1_Ukeir(Got_Focus)
'+------------------------------------+
Private Sub iH1_Ukeir_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
'+------------------------------------+
'+  iH1_Ukeir(Lost_Focus)
'+------------------------------------+
Private Sub iH1_Ukeir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Visible = False
        PB_Look.Tag = ""
    End If
End Sub
'+------------------------------------+
'+  iH1_Hinbn(Got_Focus)
'+------------------------------------+
Private Sub iH1_Hinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
'+------------------------------------+
'+  iH1_Hinbn(Lost_Focus)
'+------------------------------------+
Private Sub iH1_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Visible = False
        PB_Look.Tag = ""
    End If
End Sub
'+------------------------------------+
'+  OpH1_ZouGen(Got_Focus)
'+------------------------------------+
Private Sub OpH1_ZouGen_GotFocus(Index As Integer)
'    OpH1_ZouGen(Index).BackColor = mOPLostColor
End Sub
'+------------------------------------+
'+  OpH1_ZouGen(Lost_Focus)
'+------------------------------------+
Private Sub OpH1_ZouGen_LostFocus(Index As Integer)
    OpH1_ZouGen(Index).BackColor = mOPLostColor
End Sub
'+------------------------------------+
'+  OpH1_ZouGen(Click)
'+------------------------------------+
Private Sub OpH1_ZouGen_Click(Index As Integer)
    OpH1_ZouGen(Index).BackColor = mOPGotColor
End Sub
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   Dim w_Color     As Long
   w_Color = &H80000008 '初期色ｾｯﾄ(黒)
   Select Case SyoriKB
     Case "MNT"
          SyoriNM = "修正"
          K_Sykbnm.ForeColor = &HFF0000  '(青色)
     Case "DSP"
          SyoriNM = "照会"
          K_Sykbnm.ForeColor = &HFF00FF   '(ピンク)
     Case Else
          SyoriNM = ""
   End Select
   For Each gObj In Me.Controls
       If gObj.Name Like "i*" Then gObj.ForeColor = w_Color
       If gObj.Name Like "*Comb*" Then gObj.ForeColor = w_Color
       If gObj.Name Like "*vsGrid*" Then gObj.ForeColor = w_Color
'      * PF-Key 表示設定 *
       If SyoriKB <> "LOC" Then
          If gObj.Name = "PB_DSP" Or gObj.Name = "PB_MNT" Then
            If RTrim(Rv_Call3) = "" Then
                gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
            Else
                gObj.Visible = False
            End If
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
          PB_ENT.BtnCaption = "開始"
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
          PB_ENT.BtnCaption = "選択"
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    PB_MNT.Visible = H1Mode
    PB_DSP.Visible = H1Mode
    PB_CAN.Visible = B1Mode
    PB_Look.Visible = H1Mode
    
    If ProcHB = "H1" Then
       Lb_Kensu = ""
       PB_Look.Visible = False
    End If
    If RTrim(Rv_Call3) = "" Then
        PB_END.BtnCaption = "終了"
    Else
        PB_END.BtnCaption = "戻る"
    End If

'----------------------------------------------- 2006/03/16 Insert
    If SyoriKB = "DSP" And ProcHB = "B1" Then
       PB_ENT.Visible = False
    Else
       PB_ENT.Visible = True
    End If
'----------------------------------------------- 2006/03/16 Insert End
' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
'内示年月
    If Trim(iH1_NJYm) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    内示年月にて入力してください　　　　"
            .MB_Button = Error
            .MBOX
        End With
        iH1_NJYm.SetFocus
        Exit Function
    End If
    If Not iH1_NJYm.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    正しい内示年月を入力してください　　　"
            .MB_Button = Error
            .MBOX
        End With
        iH1_NJYm.SetFocus
        Exit Function
    End If
    
    
    If Not DBInput Then
        With CisFun
             .MB_Lines = 4
             .MB_MSG(1) = "    該当データが存在しません｡            "
             .MB_MSG(3) = "  再度抽出条件を入力してください            "
             .MB_Button = Error
             .MBOX
        End With
       Exit Function
    End If
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
'    If iB1_???? = "" Then
'       ReDim gMBMsg(5)
'       gMBMsg(2) = "  ???を入力して下さい  "
'       gMBMsg(3) = "   "
'       gMBTitle = " "
'       gMBButton = "E"
'       gRtn = MBox
'       iB1_???.SetFocus
'       Exit Function
'    End If
    Body1Chk = True
End Function
''******************************************'
''*****          分  類  表  示         *****
''******************************************'
Private Sub Op_Sort_Click(Index As Integer)
    For Each gObj In Op_Sort
        If gObj.Index = Index Then
           gObj.BackColor = mOPGotColor
           gObj.ForeColor = &H80000012
        Else
           gObj.BackColor = Frm_Sort.BackColor
           gObj.ForeColor = &HFFFFFF
        End If
    Next gObj
    If ProcHB = "H1" Then Exit Sub
    Call DBInput
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
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    DoEvents
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If SyoriKB = "MNT" Then
            Call GenMntDisp
        Else
            If RTrim(Rv_Call3) <> "" Then
                Call RtnSet
            End If
        End If
        
        ProcHB = "B1"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
    
ReturnPress_Ed:
    Call CisFrm.MousePT(1)
End Sub

'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput(Optional InType As String = "") As Boolean
    Dim wDate   As String

    DBInput = True
    On Error GoTo DBInput_Err:
    DoEvents
    
    gSL_Select = "SELECT KK.仕入先,KK.仕入先 + ':' + ISNULL(TR.略称,'') 編集仕入先,"
    gSL_Select = gSL_Select & "  KK.受入,KK.背番号,ISNULL(HN.表示品番,'') 表示品番,KK.品番,"
    gSL_Select = gSL_Select & "  CASE WHEN ISNULL(HT.変更日,'') = '' THEN HT.収容数1"
    gSL_Select = gSL_Select & "       ELSE CASE WHEN ISNULL(HT.変更日,'') <= '" & RTrim(iH1_NJYm) & "' THEN HT.収容数2"
    gSL_Select = gSL_Select & "                 ELSE HT.収容数1 END"
    gSL_Select = gSL_Select & "       END 収容数,"
    gSL_Select = gSL_Select & "  CASE WHEN ISNULL(HT.サイクル,'') <> '' THEN"
    gSL_Select = gSL_Select & "       SUBSTRING(HT.サイクル,1,1) + '-' + SUBSTRING(HT.サイクル,2,2) + '-' + SUBSTRING(HT.サイクル,4,2)"
    gSL_Select = gSL_Select & "       Else '' END 編集サイクル,"
    gSL_Select = gSL_Select & "  ISNULL(HNJ.内示数,0) 当月内示数,KK.回転枚数,KK.前月回転枚数,"
    gSL_Select = gSL_Select & "  CASE WHEN KK.増減区分 = 0 THEN KK.抜取枚数 * -1 ELSE KK.増発枚数 END 回転枚数差,"
    gSL_Select = gSL_Select & "  CASE WHEN KK.平準化区分 = 1 THEN '日'"
    gSL_Select = gSL_Select & "       WHEN KK.平準化区分 = 2 THEN '便'"
    gSL_Select = gSL_Select & "       WHEN KK.平準化区分 = 3 THEN '月'"
    gSL_Select = gSL_Select & "       Else '' END 平準化区分,"
    gSL_Select = gSL_Select & "  CASE WHEN KK.平準化区分 = 1 THEN 日上限枚数"
    gSL_Select = gSL_Select & "       WHEN KK.平準化区分 = 2 THEN 便上限枚数"
    gSL_Select = gSL_Select & "       WHEN KK.平準化区分 = 3 THEN 月上限枚数"
    gSL_Select = gSL_Select & "       ELSE 0 END 平準化枚数,"
    gSL_Select = gSL_Select & "  CASE WHEN HT.か係数 <> 0 THEN HT.か係数 ELSE NULL END か係数,HT.平準化係数"
    gSL_Select = gSL_Select & " FROM かんばん回転枚数テーブル KK"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON TR.取引先CD = KK.仕入先"
    gSL_Select = gSL_Select & "   AND TR.取引先区分 = 1"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "   ON HN.品番 = KK.品番"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HT"
    gSL_Select = gSL_Select & "   ON HT.品番 = KK.品番"
    gSL_Select = gSL_Select & "   AND HT.取引先 = KK.仕入先"
    gSL_Select = gSL_Select & "   AND HT.受入 = KK.受入"
    gSL_Select = gSL_Select & "   AND HT.品目 = 1"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN (SELECT 手配先,受入,品番,SUM(数量) 内示数 FROM 発注内示テーブル"
    gSL_Select = gSL_Select & "                   WHERE 内示年月 = '" & RTrim(iH1_NJYm) & "'"
    gSL_Select = gSL_Select & "                     AND 計画年月 = '" & RTrim(iH1_NJYm) & "'"
    gSL_Select = gSL_Select & "                  GROUP BY 手配先,受入,品番) HNJ"
    gSL_Select = gSL_Select & "   ON HNJ.手配先 = KK.仕入先"
    gSL_Select = gSL_Select & "   AND HNJ.受入 = KK.受入"
    gSL_Select = gSL_Select & "   AND HNJ.品番 = KK.品番"

    If ProcHB = "H1" And InType = "" Then
'       表示条件 セット
'       入力 Select 文作成
        mSL_SelectW = ""
'       日付
        With CisFun
            .Where_CreateKB = NewDocument
            ' 内示年月
            .Where_Create AlphaMe, "年月", iH1_NJYm
            .Where_Create AlphaMe, "KK.仕入先", iH1_Torcd
            .Where_Create AlphaMe, "KK.受入", iH1_Ukeir
            .Where_Create AlphaMe, "KK.品番", iH1_Hinbn
            mSL_SelectW = mSL_SelectW & .Where_Phrase
        End With
        If CisFun.Val2(iH1_Maisu) <> 0 Then
            mSL_SelectW = mSL_SelectW & " AND (増発枚数 >= " & CisFun.Val2(iH1_Maisu)
            mSL_SelectW = mSL_SelectW & "  OR 抜取枚数 >= " & CisFun.Val2(iH1_Maisu) & ")"
        End If
        If OpH1_ZouGen(1).Value Then
            mSL_SelectW = mSL_SelectW & " AND 増減区分 = 1"
        End If
        If OpH1_ZouGen(2).Value Then
            mSL_SelectW = mSL_SelectW & " AND 増減区分 = 0"
        End If
    End If
    
    mSL_OrderA = " ORDER BY KK.仕入先,KK.受入,KK.背番号,KK.品番"
    mSL_OrderD = " ORDER BY KK.仕入先 DESC,KK.受入 DESC,KK.背番号 DESC,KK.品番 DESC"
' 分類内容(Order By句) ｾｯﾄ
    If Op_Sort(0).Value = True Then
        gSL_Select = gSL_Select & mSL_SelectW & mSL_OrderA
    Else
        gSL_Select = gSL_Select & mSL_SelectW & mSL_OrderD
    End If
    If Not DBGridDisp Then
       DBInput = False
       iH1_Torcd.SetFocus
       GoTo DBInput_Ed
    End If
    On Error Resume Next
    MaxCnt = Adodc1.Recordset.RecordCount
    If MaxCnt = 0 Then
       DBInput = False
       GoTo DBInput_Ed
    End If
    Lb_Kensu.Visible = True
    Call TotalDisp
    On Error GoTo 0
DBInput_Ed:
    Call CisFrm.MousePT(0)
    SendKeys "{END}"
    SendKeys "{RIGHT}"
    SendKeys "{LEFT}"
    SendKeys "{HOME}"
    Exit Function
DBInput_Err:
    CisFun.ErrorBox
    Call CisFrm.MousePT(0)
End Function
'+----------------------------------------------------------+
'+                                                          +
'+                                                          +
'+      グリッド処理(TrueDBGrid)                             +
'+                                                          +
'+                                                          +
'+----------------------------------------------------------+
'+----------------------------------------------------------+
'+      ADODC
'+----------------------------------------------------------+
Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Adodc1.RecordSource <> "" Then Call TotalDisp
End Sub
'+----------------------------------------------------------+
'+      ADODC1(件数表示)                                     +
'+----------------------------------------------------------+
Private Sub TotalDisp()
    NowCnt = Adodc1.Recordset.AbsolutePosition
    MaxCnt = Adodc1.Recordset.RecordCount
    If NowCnt < 0 Then Exit Sub
    Lb_Kensu.Caption = Format(NowCnt, "###,##0") & " / " & _
                       Format(MaxCnt, "###,##0")
End Sub
'+----------------------------------------------------------+
'+      グリッド    《 MouseDown  イベント   》               +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If DBGrid1.ColContaining(x) < 0 Or DBGrid1.RowContaining(Y) < 0 Then Exit Sub
    DBGrid1.Row = DBGrid1.RowContaining(Y)
End Sub

Private Sub DBGrid1_DblClick()
    If SyoriKB = "DSP" And RTrim(Rv_Call3) <> "" Then
        Call RtnSet
        Exit Sub
    End If
    If SyoriKB = "MNT" Then
        Call GenMntDisp
    End If
End Sub

''+----------------------------------------------------------+
''+      グリッド    《 MouseUp  イベント   》                +
''+----------------------------------------------------------+
'Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'' ｴﾗｰ内容表示
'    If DBGrid1.RowContaining(Y) >= 0 Then
'        Call ErrDisp
'    End If
'End Sub
'Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeySpace Then
'    ' ｴﾗｰ内容表示
'        Call ErrDisp
'    End If
'End Sub
'+-----------------------+
'+  グリッド内容初期化    +
'+-----------------------+
Private Sub DBGridClr()
    Adodc1.Recordset.Close
    Adodc1.RecordSource = ""
    DBGrid1.Close True
    Lb_Kensu.Caption = ""
End Sub
'+---------------------+
'+  グリッド内容表示    +
'+---------------------+
Private Function DBGridDisp()
    DBGridDisp = False
    On Error GoTo DBGridDisp_ER:
    If Me.MousePointer <> 11 Then
       Me.MousePointer = 11
    End If
    On Error Resume Next
    Adodc1.RecordSource = gSL_Select
    DBGrid1.ReBind
    Adodc1.Refresh
    DBGrid1.Refresh
    With DBGrid1
         .SetFocus
    End With
    SendKeys "{HOME}"
    If Me.MousePointer <> 1 Then
       Me.MousePointer = 1
    End If
    DBGridDisp = True
    On Error GoTo 0
    Exit Function
DBGridDisp_ER:
    If Err <> 40002 Then        ' RDO タイムアウト
       CisFun.ErrorBox
    End If
End Function
'+-----------------------------------+
'+    増減・平準化メンテ画面表示       +
'+-----------------------------------+
Private Sub GenMntDisp()
    Static wErrCode As String
'---------------------------------- 2006/03/16 Delete
'    With Adodc1
'        RV_Torcd = .Recordset("仕入先")
'        RV_Ukeir = .Recordset("受入")
'        RV_Seban = .Recordset("背番号")
'        RV_Hinbn = .Recordset("品番")
'    End With
'    Rv_Call2 = "M"
'
'    CKH0110.Show vbModal
'    Unload CKH0110
'    Set CKH0110 = Nothing
'
'---------------------------------- 2006/03/16 Delete End
'
'
'---------------------------------- 2006/03/16 Insert
    RV_Left = 0
    RV_Top = 0
    RV_YM = RTrim(iH1_NJYm)

    With Adodc1
        RV_Torcd = .Recordset("仕入先")
        RV_Ukeir = .Recordset("受入")
        RV_Seban = .Recordset("背番号")
        RV_Hinbn = .Recordset("品番")
    End With
    '開始日の検索、存在しなければ基本開始日をセット
    gSL_Select = "SELECT 1 区分,* FROM かんばん増減開始日マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(RV_Torcd) & "'"
    gSL_Select = gSL_Select & " AND   受入 = ''"
    gSL_Select = gSL_Select & " AND   品番 = ''"
    gSL_Select = gSL_Select & " UNION "
    gSL_Select = gSL_Select & " SELECT 2 区分,* FROM かんばん増減開始日マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(RV_Torcd) & "'"
    gSL_Select = gSL_Select & " AND   受入 = '" & RTrim(RV_Ukeir) & "'"
    gSL_Select = gSL_Select & " AND   品番 = ''"
    gSL_Select = gSL_Select & " UNION "
    gSL_Select = gSL_Select & " SELECT 3 区分,* FROM かんばん増減開始日マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(RV_Torcd) & "'"
    gSL_Select = gSL_Select & " AND   受入 = '" & RTrim(RV_Ukeir) & "'"
    gSL_Select = gSL_Select & " AND   品番 = '" & RTrim(RV_Hinbn) & "'"
    gSL_Select = gSL_Select & " ORDER BY 区分 DESC"
    If KZSRead(gSL_Select, 1, 1) Then
       mStartDay = CisFun.Val2(KZS.開始日)
    Else
       mStartDay = mBaseStartDay
    End If
    
    If mStartDay <> 1 Then
       mStartYMD = RTrim(iH1_NJYm) & "01"
       CisFun.Date_RsFormat = fYMD
       CisFun.Date_Add mStartYMD, D, mStartDay
       mStartYMD = CisFun.Date_Result
    Else
       mStartYMD = RTrim(iH1_NJYm) & "01"
    End If
    With CisFun
        .Date_RsFormat = fYMD
        .Date_End RTrim(iH1_NJYm) & "01"
        mEndYMD = .Date_Result
    End With
    If mStartDay <> 1 Then
       With CisFun
           .Date_RsFormat = fYMD
           .Date_Add mEndYMD, D, mStartDay
           mEndYMD = .Date_Result
       End With
    End If
    RV_SYmd = mStartYMD
    RV_EYmd = mEndYMD
    
    Rv_Call2 = "X"
    
    CKH0170.Show vbModal
    Unload CKH0170
    Set CKH0170 = Nothing
'---------------------------------- 2006/03/16 Insert End

    ProcHB = "B1"
    Call DispChange(ProcHB)
End Sub
'+-------------------------------------+
'+      選択内容ｾｯﾄ & 復帰ｽﾃｰﾀｽｾｯﾄ      +
'+-------------------------------------+
Private Sub RtnSet()
    With Adodc1
        RV_Torcd = .Recordset("仕入先")
        RV_Ukeir = .Recordset("受入")
        RV_Seban = .Recordset("背番号")
        RV_Hinbn = .Recordset("品番")
    End With
    Adodc1.Recordset.Close
    RV_Rtn = True
    Call PB_END_Click
End Sub



