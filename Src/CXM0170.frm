VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Begin VB.Form CXM0170 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "型費データ入力"
   ClientHeight    =   11025
   ClientLeft      =   1035
   ClientTop       =   1740
   ClientWidth     =   13695
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
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
   ScaleHeight     =   11025
   ScaleWidth      =   13695
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   405
      Index           =   0
      Left            =   10815
      ScaleHeight     =   345
      ScaleWidth      =   2325
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2550
      Width           =   2385
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   345
         Index           =   1
         Left            =   15
         Top             =   0
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   609
         Caption         =   " 登録済件数"
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
         Begin Cis3D_v60.CIS3D lb_Kensu 
            Height          =   270
            Left            =   1290
            Top             =   30
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   476
            BackColor       =   16777152
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -60
         End
      End
   End
   Begin VB.Timer Timer4 
      Left            =   210
      Top             =   810
   End
   Begin VB.PictureBox B2_Area1 
      Height          =   4590
      Left            =   1485
      ScaleHeight     =   4530
      ScaleWidth      =   11655
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4080
      Width           =   11715
      Begin vsFlexLib.vsFlexArray VsGrid1 
         Height          =   4380
         Left            =   90
         TabIndex        =   8
         Top             =   90
         Width           =   11505
         _Version        =   196608
         _ExtentX        =   20294
         _ExtentY        =   7726
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   "No|適用開始年月|適用終了年月|月発生金額    |回数 |投 資 金 額    |切替年月    ||||||||||"
         Rows            =   13
         Cols            =   17
         BackColor       =   12648447
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   330
         SelectionMode   =   1
         AllowUserResizing=   1
      End
   End
   Begin VB.PictureBox B1_Area1 
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
      Height          =   1020
      Left            =   1500
      ScaleHeight     =   960
      ScaleWidth      =   11640
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2955
      Width           =   11700
      Begin Cis3D_v60.CIS3D Back_TSYm 
         Height          =   885
         Left            =   60
         Top             =   45
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "適用開始年月"
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
         Begin CisYM_V60.CisYM iB1_TSYm 
            Height          =   375
            Left            =   135
            TabIndex        =   2
            Top             =   360
            Width           =   1530
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
      Begin Cis3D_v60.CIS3D Back_Kaisu 
         Height          =   885
         Left            =   5340
         Top             =   45
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "回数"
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
         Begin CisText_V60.CisText iB1_Kaisu 
            Height          =   360
            Left            =   165
            TabIndex        =   5
            Top             =   360
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
      Begin Cis3D_v60.CIS3D Back_ToKin 
         Height          =   885
         Left            =   7875
         Top             =   45
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "投資金額"
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
         Begin CisText_V60.CisText iB1_ToKin 
            Height          =   360
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   635
            cFormat         =   "#,###,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "#,###,###,###"
            MaxLength       =   13
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "##########"
            cILength        =   10
         End
      End
      Begin Cis3D_v60.CIS3D Back_MmKin 
         Height          =   885
         Left            =   3600
         Top             =   45
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "月発生金額"
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
         Begin CisText_V60.CisText iB1_MmKin 
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   661
            cFormat         =   "###,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "###,###,###"
            MaxLength       =   11
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "#########"
            cILength        =   9
         End
      End
      Begin Cis3D_v60.CIS3D Back_TEYm 
         Height          =   885
         Left            =   1830
         Top             =   45
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "適用終了年月"
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
         Begin CisYM_V60.CisYM iB1_TEYm 
            Height          =   375
            Left            =   135
            TabIndex        =   3
            Top             =   360
            Width           =   1530
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
      Begin Cis3D_v60.CIS3D Back_KYm 
         Height          =   885
         Left            =   9855
         Top             =   45
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "切替年月"
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
         Begin CisYM_V60.CisYM iB1_KYm 
            Height          =   375
            Left            =   135
            TabIndex        =   7
            Top             =   360
            Width           =   1530
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
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   885
         Left            =   6060
         Top             =   45
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1561
         ForeColor       =   0
         Caption         =   "最終月調整金額"
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
         Begin Cis3D_v60.CIS3D B1lb_CKingk 
            Height          =   315
            Left            =   165
            Top             =   360
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            BackColor       =   16777152
            Caption         =   "###,###,###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -60
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
      Left            =   1515
      ScaleHeight     =   870
      ScaleWidth      =   7245
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1590
      Width           =   7305
      Begin Cis3D_v60.CIS3D HLB_Hinnm 
         Height          =   795
         Left            =   60
         Top             =   45
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "契 約 先"
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
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1125
            Top             =   360
            Width           =   2505
            _ExtentX        =   4419
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   135
            TabIndex        =   0
            Top             =   360
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
      End
      Begin Cis3D_v60.CIS3D BLB_Zaist 
         Height          =   795
         Left            =   3825
         Top             =   45
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "品　番"
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
         cPositionX      =   90
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   360
            Left            =   105
            TabIndex        =   1
            Top             =   360
            Width           =   3210
            _ExtentX        =   5662
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
         End
      End
   End
   Begin VB.Timer Timer3 
      Left            =   1980
      Top             =   810
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
      Height          =   1365
      Left            =   9120
      ScaleHeight     =   1305
      ScaleWidth      =   4365
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1125
      Width           =   4425
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1755
         Top             =   660
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
         Left            =   2625
         Top             =   660
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
         Left            =   3480
         Top             =   660
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   2610
         Top             =   30
         Visible         =   0   'False
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
      Begin CisBtn_60.CisBtn PB_NEXT 
         Height          =   585
         Left            =   900
         Top             =   660
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
      Begin CisBtn_60.CisBtn PB_BACK 
         Height          =   585
         Left            =   45
         Top             =   660
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   1740
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
      Begin CisBtn_60.CisBtn PB_MNT 
         Height          =   585
         Left            =   885
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
   End
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   990
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   540
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Left            =   930
      Top             =   660
   End
   Begin VB.Timer Timer2 
      Left            =   1440
      Top             =   720
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
      Caption         =   "【　型費データ入力  】"
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
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'なし
         Height          =   345
         Index           =   1
         Left            =   10620
         ScaleHeight     =   345
         ScaleWidth      =   3675
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   30
         Width           =   3675
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
               Top             =   30
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
            Left            =   0
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
               Top             =   30
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
            Left            =   1230
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
               Top             =   30
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
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14400
         Top             =   90
         Width           =   885
         _ExtentX        =   1561
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
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   285
         Left            =   8070
         Top             =   60
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   503
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
Attribute VB_Name = "CXM0170"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************************'
'**
'**       フォーム名    :   型費データ入力
'**       フォームID    :   CXM0170
'**       処理概要      :
'**
'**       作  成  日    :   2004/03/29  By CIS SIMPR-A
'**       変  更  日    :   2005/04/07  By CIS  ・[前進][後退]時、契約先名称の再読込みを追加。
'**                                             ・修正時、対象データが消えないように条件を変更。
'**       変  更  日    :
'**
'**
'********************************************************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)
    Dim CisVsGrid1      As New CisVsGrid3
    Dim SaveSelect      As String
    
    Dim SaveRenban      As Long          ' 連番 退避
    Dim StartYyMm       As Long          ' 開始年月 退避
    
    Dim SaveUYm         As String       ' 売上処理年月
    Dim SaveRow         As Integer
    Dim SaveKeisan      As Byte
    Dim RowColFlg       As Boolean

    Dim mGotColor       As Long
    Dim mLostColor      As Long
    
    Dim KDouble         As Double
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
'
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
'
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
'
    Dummy.Left = -1000
    Dummy.Enabled = False
'    Dummy2.Left = -1000
'    Dummy3.Left = -1000
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔(基本)
    Timer2.Enabled = False   ' 処理区分点滅禁止(基本)
'    Timer3.Interval = 500    ' 処理区分点滅間隔(取引先)
'    Timer3.Enabled = False   ' 処理区分点滅禁止(取引先)
'
    '+--------------------------------+
    '+  グリッド初期設定
    '+--------------------------------+
    With CisVsGrid1
        Set .GridObj = VsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 350
        .PatanMax = 1
        .SelectCol = 12
        .InitGet = False
        .Init
    End With
    Call GridInit1
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = ""
    Call SyoriName(SyoriKB)
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)
'
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
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
           Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click    '【新規】
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    '【修正】
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
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
        Call CisVsGrid1.vsColor(消去個別)
        Call HeadBodyClear("B1")
        If Val(lb_Kensu) <> 0 Then
            ProcHB = "B2"
            SyoriKB = "MNT"
            Call VSGrid1_Click
        Else
            ProcHB = "H1"
            SyoriKB = ""
        End If
        Call SyoriName(SyoriKB)
        Call DispChange(ProcHB)
        GoTo PB_CAN_Ed
    End If
            
    If ProcHB = "B2" Then
        SyoriKB = ""
        Call SyoriName(SyoriKB)
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        GoTo PB_CAN_Ed
    End If

PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    新　規  キ ー(F1)        +
'+----------------------------+
Private Sub PB_ADD_Click()
    
    If SyoriKB <> "ADD" Then
        ProcHB = "B1"
        Call DispChange(ProcHB)
        SyoriKB = "ADD"
        Call SyoriName(SyoriKB)
        Call HeadBodyClear("B1")
        RowColFlg = True
        Call CisVsGrid1.vsColor(消去個別)
        RowColFlg = False
        
        iB1_TSYm.SetFocus
    End If
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)        +
'+----------------------------+
Private Sub PB_MNT_Click()
    
    If SyoriKB <> "MNT" Then
        ProcHB = "B2"
        Call DispChange(ProcHB)
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
        RowColFlg = True
        Call CisVsGrid1.vsColor(消去個別)
        RowColFlg = False

        
        VsGrid1.SetFocus
        Call VSGrid1_Click
    End If
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)        +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
        ProcHB = "B2"
        Call DispChange(ProcHB)
        SyoriKB = "DEL"
        Call SyoriName(SyoriKB)
        RowColFlg = True
        Call CisVsGrid1.vsColor(消去個別)
        RowColFlg = False
        
        VsGrid1.SetFocus
        Call VSGrid1_Click
    End If

End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    
    Dim ReadKey     As String * 32
    ReadKey = Space(32)
    Mid(ReadKey, 1, 7) = RTrim(iH1_Torcd)
    Mid(ReadKey, 8, 25) = RTrim(iH1_Hinbn)
    
    gSL_Select = " SELECT 取引先,品番"
    gSL_Select = gSL_Select & " FROM 型費テーブル "
    gSL_Select = gSL_Select & " WHERE CAST(取引先 AS CHAR(7)) + 品番 < '" & RTrim(ReadKey) & "'"
    gSL_Select = gSL_Select & " ORDER BY 取引先 DESC,品番 DESC "
'*-------------------------------------------------▽▽ 2005.04.07 Upd Start ▽▽------------------------------*
'    If KTTRead(gSL_Select, 1) Then
'        iH1_Torcd = RTrim(KTT.取引先)
'        iH1_Hinbn = RTrim(KTT.品番)
'    Else
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "     先頭レコードです            "
'            .MB_Title = "【後退】"
'            .MB_Button = OK
'            .MBOX
'        End With
'        GoTo PB_Back_Ed
'    End If
''
'    Call DBInput("Back")
'
'PB_Back_Ed:
'    iH1_Hinbn.SetFocus
'    Call CisFrm.MousePT(1)
'*
    If Not KTTRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     先頭レコードです            "
            .MB_Title = "【後退】"
            .MB_Button = OK
            .MBOX
        End With
        Call CisFrm.MousePT(1)
        Exit Sub
    End If
    iH1_Torcd = RTrim(KTT.取引先)
    iH1_Hinbn = RTrim(KTT.品番)
    Call iH1_Torcd_LostFocus
    
    Call DBInput("Back")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Hinbn.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
'*-------------------------------------------------△△ 2005.04.07 Upd End △△--------------------------------*
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    Dim ReadKey     As String * 32
    ReadKey = Space(32)
    Mid(ReadKey, 1, 7) = RTrim(iH1_Torcd)
    Mid(ReadKey, 8, 25) = RTrim(iH1_Hinbn)
    
    gSL_Select = " SELECT 取引先,品番"
    gSL_Select = gSL_Select & " FROM 型費テーブル "
    gSL_Select = gSL_Select & " WHERE CAST(取引先 AS CHAR(7)) + 品番 > '" & RTrim(ReadKey) & "'"
    gSL_Select = gSL_Select & " ORDER BY 取引先,品番 "
'*-------------------------------------------------▽▽ 2005.04.07 Upd Start ▽▽------------------------------*
'    If KTTRead(gSL_Select, 1) Then
'        iH1_Torcd = RTrim(KTT.取引先)
'        iH1_Hinbn = RTrim(KTT.品番)
'    Else
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "     最終レコードです            "
'            .MB_Title = "【前進】"
'            .MB_Button = OK
'            .MBOX
'        End With
'        GoTo PB_Next_Ed
'    End If
''
'    Call DBInput("Next")
'
'PB_Next_Ed:
'    iH1_Hinbn.SetFocus
'    Call CisFrm.MousePT(1)
'*
    If Not KTTRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     最終レコードです            "
            .MB_Title = "【前進】"
            .MB_Button = OK
            .MBOX
        End With
        Call CisFrm.MousePT(1)
        Exit Sub
    End If

    iH1_Torcd = RTrim(KTT.取引先)
    iH1_Hinbn = RTrim(KTT.品番)
    Call iH1_Torcd_LostFocus
    
    Call DBInput("Next")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Hinbn.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
'*-------------------------------------------------△△ 2005.04.07 Upd End △△--------------------------------*
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub

End Sub
'+----------------------------+
'+       契約先検索            +
'+----------------------------+
Private Sub Look_Torcd()

    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 0
    RV_TorKb1 = 1
    
    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RTrim(RV_TorcdK)
        Call iH1_Torcd_LostFocus
    End If
End Sub
'+----------------------------+
'+       品番検索              +
'+----------------------------+
Private Sub Look_Hinbn()

    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 0
    RVI_SK = 9
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iH1_Hinbn = RTrim(RV_Hinbn)
    End If
End Sub
'+---( H1 )----------------------------------------------------------------------------
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
    If TorNmGet(iH1_Torcd, 2) Then
        H1lb_Tornm = TRM.略称
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
'+---( B1 )----------------------------------------------------------------------------
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
'+------------------------------------------------------------------+
'+      基本情報
'+------------------------------------------------------------------+
Private Sub SyoriName(SYKB As String)
    Dim w_Color     As Long
    w_Color = &H80000008 '初期色ｾｯﾄ(黒)
   
'   H2LB_Size_B.Enabled = False
   
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
'      * PF-Key 表示設定 *
        If SyoriKB <> "LOC" Then
           If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
              gObj.Name = "PB_DEL" Then
              gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
           End If
        End If
    Next gObj
   
    If SyoriKB = "ADD" Then
        Back_TSYm.Enabled = True
        Back_TEYm.Enabled = True
        Back_MmKin.Enabled = True
        Back_Kaisu.Enabled = True
        Back_ToKin.Enabled = True
    Else
        Back_TSYm.Enabled = False
        Back_TEYm.Enabled = False
        Back_MmKin.Enabled = False
        Back_Kaisu.Enabled = False
        Back_ToKin.Enabled = False
    End If
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
    

    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    PB_ADD.Visible = Not H1Mode
    PB_MNT.Visible = Not H1Mode
    PB_DEL.Visible = Not H1Mode
    
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
    
    On Error GoTo HeadBodyClear_ER
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then
                .Text = ""
              '  .Enabled = True
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
                .BackColor = mLostColor
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    
    
    If HeadBody = "B2" Or HeadBody = "B" Or HeadBody = "ALL" Then
        CisVsGrid1.Clear
        Call CisVsGrid1.Clear
        lb_Kensu = ""
    End If
    Exit Sub
HeadBodyClear_ER:
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        If ProcHB = "B2" Then
            Call VSGrid1_Click
            VsGrid1.SetFocus
        Else
            iB1_TSYm.SetFocus
        End If
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut(0) Then
            GoTo ReturnPress_Ed
        End If
        If SyoriKB = "ADD" Then
            ProcHB = "B1"
            Call DispChange(ProcHB)
            iB1_TSYm.SetFocus
        Else
            ProcHB = "B2"
            Call DispChange(ProcHB)
            Call VSGrid1_Click
            VsGrid1.SetFocus
        End If
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "B2" Then
        If Not Body2Chk Then GoTo ReturnPress_Ed
' << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then
                RowColFlg = True
                Call CisVsGrid1.vsColor(消去個別)
                RowColFlg = False
                GoTo ReturnPress_Ed
            End If
            Call HeadBodyClear("B1")
            If CisFun.Val2(lb_Kensu) <> 0 Then
                ProcHB = "B2"
                Call DispChange(ProcHB)
                Call SyoriName(SyoriKB)
                Call VSGrid1_Click
                Call VsGrid1.SetFocus
            Else
                ProcHB = "B1"
                Call DispChange(ProcHB)
                SyoriKB = "ADD"
                Call SyoriName(SyoriKB)
                iB1_TSYm.SetFocus
            End If
        Else
            ProcHB = "B1"
            Call DispChange(ProcHB)
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
'
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
' 契約先
    If Trim(iH1_Torcd) = "" Then
        CisFun.MB_MSG(2) = "    契約先を入力して下さい            "
        CisFun.MBOX
       iH1_Torcd.SetFocus
       Exit Function
    End If
    If Not TorNmGet(iH1_Torcd, 2) Then
        CisFun.MB_MSG(2) = "     契約先 未登録！            "
        CisFun.MB_MSG(4) = "    （取引先マスタ）            "
        CisFun.MBOX
       iH1_Torcd.SetFocus
       Exit Function
    End If
    SaveKeisan = TRM.計算単位
' 品番
    If Trim(iH1_Hinbn) = "" Then
        CisFun.MB_MSG(2) = "    品番を入力して下さい            "
        CisFun.MBOX
       iH1_Hinbn.SetFocus
       Exit Function
    End If
' 品番(全桁)の存在
    gSL_Select = "SELECT COUNT(品番) 件数 FROM 品番取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   AND 契約先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND 品目 = 0"
    If HTMRead(gSL_Select, 1) Then
        If HTM.件数 = 0 Then
            ' 品番(10桁)の存在
            gSL_Select = "SELECT COUNT(品番) 件数 FROM 品番取引先マスタ"
            gSL_Select = gSL_Select & " WHERE SUBSTRING(品番,1,10) = '" & RTrim(iH1_Hinbn) & "'"
            gSL_Select = gSL_Select & "   AND 契約先 = '" & RTrim(iH1_Torcd) & "'"
            gSL_Select = gSL_Select & "   AND 品目 = 0"
            If HTMRead(gSL_Select, 1) Then
                If HTM.件数 = 0 Then
                    CisFun.MB_MSG(1) = "    品番 未登録！(品番取引先マスタ)            "
                    CisFun.MB_MSG(3) = "    続行しますか？            "
                    CisFun.MB_Button = OK_CAN
                    If Not CisFun.MBOX Then
                        iH1_Hinbn.SetFocus
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
'型費テーブル
    gSL_Select = "SELECT 取引先,品番,連番,適用開始年月,適用終了年月,"
    gSL_Select = gSL_Select & " 月発生金額,調整金額,投資金額,回数,切替年月, "
    gSL_Select = gSL_Select & " MIN(対象開始年月)対象開始年月,MAX(実績年月)実績年月,"
    gSL_Select = gSL_Select & " MAX(予想年月)予想年月,MIN(金額)金額"
    gSL_Select = gSL_Select & " FROM 型費テーブル"
    gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " GROUP BY 取引先,品番,連番,適用開始年月,適用終了年月,"
    gSL_Select = gSL_Select & " 月発生金額,調整金額,投資金額,回数,切替年月 "
    gSL_Select = gSL_Select & " ORDER BY 適用開始年月"
    SaveSelect = gSL_Select
    If Not DBInput("INV") Then
        SyoriKB = "ADD"
        ProcHB = "B1"
    Else
        SyoriKB = "MNT"
        ProcHB = "B2"
    End If
    
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    
    Body1Chk = False
        
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
    ' 適用開始年月
    If Trim(iB1_TSYm) = "" Then
        CisFun.MB_MSG(2) = "    適用開始年月を入力して下さい            "
        CisFun.MBOX
       iB1_TSYm.SetFocus
       Exit Function
    End If
    If Not iB1_TSYm.cChkResult Then
        CisFun.MB_MSG(2) = "    正しい適用開始年月を入力して下さい            "
        CisFun.MBOX
       iB1_TSYm.SetFocus
       Exit Function
    End If
    ' 適用終了年月
    If Trim(iB1_TEYm) = "" Then
        CisFun.MB_MSG(2) = "    適用終了年月を入力して下さい            "
        CisFun.MBOX
       iB1_TEYm.SetFocus
       Exit Function
    End If
    If Not iB1_TEYm.cChkResult Then
        CisFun.MB_MSG(2) = "    正しい適用終了年月を入力して下さい            "
        CisFun.MBOX
       iB1_TEYm.SetFocus
       Exit Function
    End If
    ' 適用開始年月 > 適用終了年月
    If Trim(iB1_TSYm) > Trim(iB1_TEYm) Then
        CisFun.MB_MSG(2) = "    適用開始年月 > 適用終了年月 ！           "
        CisFun.MBOX
       iB1_TEYm.SetFocus
       Exit Function
    End If
    
    ' 月発生金額
    If CisFun.Val2(iB1_MmKin) = 0 Then
        CisFun.MB_MSG(2) = "     月発生金額を入力して下さい！          "
        CisFun.MBOX
        iB1_MmKin.SetFocus
        Exit Function
    End If
    
    ' 回数
    gCnt1 = DateDiff("m", CDate(iB1_TSYm.cYear & "/" & iB1_TSYm.cMonth & "/01"), CDate(iB1_TEYm.cYear & "/" & iB1_TEYm.cMonth & "/01"))
    If CisFun.Val2(iB1_ToKin) = 0 Then
        gCnt1 = gCnt1 + 1
    Else
        If CisFun.Val2(iB1_MmKin) * (gCnt1 + 1) = CisFun.Val2(iB1_ToKin) Then
            gCnt1 = gCnt1 + 1
        End If
    End If
    If CisFun.Val2(iB1_Kaisu) = 0 Then
        iB1_Kaisu = Format(gCnt1, "##")
    End If
    B1lb_CKingk = ""
    If CisFun.Val2(iB1_ToKin) <> 0 And CisFun.Val2(iB1_ToKin) - (CisFun.Val2(iB1_MmKin) * gCnt1) <> 0 Then
        B1lb_CKingk = Format(CisFun.Val2(iB1_ToKin) - (CisFun.Val2(iB1_MmKin) * gCnt1), "#,###")
    End If
    If gCnt1 <> CisFun.Val2(iB1_Kaisu) Then
        CisFun.MB_MSG(2) = "     回数に誤りがあります！          "
        CisFun.MBOX
        iB1_Kaisu.SetFocus
        Exit Function
    End If
'投資金額
    If CisFun.Val2(iB1_ToKin) > 2000000000 Then
        CisFun.MB_MSG(2) = "     投資金額は規定値オーバー！(2,000,000,000迄)          "
        CisFun.MBOX
        iB1_ToKin.SetFocus
        Exit Function
    End If
    If CisFun.Val2(B1lb_CKingk) < 0 Then
        CisFun.MB_MSG(2) = "     投資金額 異常          "
        CisFun.MBOX
        iB1_ToKin.SetFocus
        Exit Function
    End If
    
    ' 切替年月
    If Trim(iB1_KYm) <> "" Then
        If Not iB1_KYm.cChkResult Then
            CisFun.MB_MSG(2) = "    正しい切替年月を入力して下さい            "
            CisFun.MBOX
           iB1_KYm.SetFocus
           Exit Function
        End If
        If Trim(iB1_KYm) > Trim(iB1_TEYm) Then
            CisFun.MB_MSG(2) = "    切替年月 > 適用終了年月 ！           "
            CisFun.MBOX
           iB1_KYm.SetFocus
           Exit Function
        End If
    End If
    
    Body1Chk = True
End Function
'****************************
'*      Body2ﾁｪｯｸ処理
'****************************
Private Function Body2Chk() As Boolean
    
    Body2Chk = False
        
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
    ' 選択
    If CisFun.Val2(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid1.FixedGet(8))) = 0 Then Exit Function
    
    RowColFlg = True
    If SyoriKB = "MNT" Then
        Call CisVsGrid1.vsColor(個別, &HC0FFC0)
    Else
        Call CisVsGrid1.vsColor(個別, &HFFC0FF)
    End If
    RowColFlg = False
    
    
    Body2Chk = True
End Function
'***********************************************************************************
'*
'*       データベース入力 ＆ 表示
'*
'***********************************************************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    Call HeadBodyClear("B")
    
    
    gSL_Select = "SELECT 取引先,品番,連番,適用開始年月,適用終了年月,"
    gSL_Select = gSL_Select & " 月発生金額,調整金額,投資金額,回数,切替年月, "
    gSL_Select = gSL_Select & " MIN(対象開始年月)対象開始年月,MAX(実績年月)実績年月,"
    gSL_Select = gSL_Select & " MAX(予想年月)予想年月,MIN(金額)金額"
    gSL_Select = gSL_Select & " FROM 型費テーブル"
    gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " GROUP BY 取引先,品番,連番,適用開始年月,適用終了年月,"
    gSL_Select = gSL_Select & " 月発生金額,調整金額,投資金額,回数,切替年月 "
    gSL_Select = gSL_Select & " ORDER BY 適用開始年月"
    SaveSelect = gSL_Select
    If Not KTTRead(gSL_Select, 0) Then
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
        Call Grid1Set
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
'+----------------------------+
'+      削 除 処 理            +
'+----------------------------+
Private Function DBDelete()
    DBDelete = False
'  ( 確認メッセージ )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    gSL_Select = "SELECT COUNT(*) 件数 FROM 型費テーブル "
    gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   AND 連番 = " & CisFun.Val2(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid1.FixedGet(8))) & ""
    gSL_Select = gSL_Select & "   AND 実績年月 <> ''"
    Call KTTRead(gSL_Select, 1)
    With CisFun
        .MB_Lines = 5
        .MB_MSG(1) = "     削除処理を行います      "
        If KTT.件数 <> 0 Then
            .MB_MSG(3) = "     売上実績計上済データが存在します。      "
        End If
        .MB_MSG(5) = "     よろしいですか？      "
        .MB_Title = "削除処理"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    With CisDB
        On Error GoTo DBDelete_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 削除処理 )
        .SQL = "DELETE FROM 型費テーブル "
        .SQL = .SQL & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
        .SQL = .SQL & "   AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
        .SQL = .SQL & "   AND 連番 IN (" & CisFun.Val2(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid1.FixedGet(8))) & ","
        .SQL = .SQL & CisFun.Val2(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid1.FixedGet(8))) + 100 & ")"
        .DBExec

        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        
        iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")
        
        On Error GoTo 0
    End With
    
'Grid再表示
    Call vsGrid1_ReSetDisp
    
    DBDelete = True
    Exit Function
DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut(KB As Byte)
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

'
    SaveRenban = 0
    If SyoriKB = "ADD" Then
        gSL_Select = "SELECT MAX(連番)連番 FROM 型費テーブル"
        gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
        If KTTRead(gSL_Select, 1) Then
            SaveRenban = KTT.連番
        End If
        SaveRenban = SaveRenban + 1
    Else
        SaveRenban = CisFun.Val2(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid1.FixedGet(8)))
    End If
'
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    On Error GoTo DBPut_Err
    If SyoriKB = "ADD" Then
    '*=================*
    '*  ( 追加処理 )
    '*=================*
        StartYyMm = CisFun.Val2(iB1_TSYm)
        Call KTTInsert_Proc
    
        iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
    Else
    '*=================*
    '*  ( 更新処理 )
    '*=================*
    ' 型費テーブル　削除
        gSL_Select = " DELETE FROM 型費テーブル"
        gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & "   AND 連番 = " & SaveRenban & ""
        gSL_Select = gSL_Select & "   AND 実績年月 = ''"
        gSL_Select = gSL_Select & "   AND 予想年月 = ''"
        CisDB.SQL = gSL_Select
        CisDB.DBExec
    ' 型費テーブル（赤）　削除
        gSL_Select = " DELETE FROM 型費テーブル"
        gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & "   AND 連番 = " & SaveRenban + 100 & ""
        gSL_Select = gSL_Select & "   AND 実績年月 = ''"
        gSL_Select = gSL_Select & "   AND 予想年月 = ''"
        CisDB.SQL = gSL_Select
        CisDB.DBExec
'*    ' 切替日　更新                        '2005.04.07 Del
'*        Call KTTUpdate_Proc               '2005.04.07 Del Del&Insしている為、無効の処理
    ' 型費　再作成
        Call KTTInsertMnt_Proc
'
        iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
    End If
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
'''    Else
'''' ( 更新処理 )
''''        KTT.更新者 = gTanto
'''        Call KTTUpdate_Proc
'''        iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
'''    End If
''''

'Grid再表示
    Call vsGrid1_ReSetDisp
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
Public Function KTTInsert_Proc() As Boolean
    KTTInsert_Proc = False
    
    Call ItemsClearKTT
    With KTT
        .取引先 = RTrim(iH1_Torcd)
        .品番 = RTrim(iH1_Hinbn)
        .連番 = SaveRenban
        .適用開始年月 = RTrim(iB1_TSYm)
        .適用終了年月 = RTrim(iB1_TEYm)
        .月発生金額 = CisFun.Val2(iB1_MmKin)
        .投資金額 = CisFun.Val2(iB1_ToKin)
        .回数 = RTrim(iB1_Kaisu)
        .調整金額 = CisFun.Val2(B1lb_CKingk)
        .切替年月 = RTrim(iB1_KYm)
        
        gCnt2 = 0
        For gCnt1 = StartYyMm To CisFun.Val2(iB1_TEYm)
            If gCnt1 Mod 100 > 12 Then
                gCnt1 = gCnt1 + 88
            End If
            .対象開始年月 = Format(gCnt1, "000000")
            gCnt2 = gCnt2 + 1
            If gCnt2 > CisFun.Val2(iB1_Kaisu) Then
                .金額 = CisFun.Val2(iB1_ToKin) - (CisFun.Val2(iB1_MmKin) * CisFun.Val2(iB1_Kaisu))
            Else
                .金額 = CisFun.Val2(iB1_MmKin)
            End If
            .作成者 = gTanto
                
            gSL_Select = " SELECT 対象開始年月 FROM 型費テーブル"
            gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
            gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
            gSL_Select = gSL_Select & "   AND 連番 = " & SaveRenban & ""
            gSL_Select = gSL_Select & "   AND 対象開始年月 = '" & RTrim(.対象開始年月) & "'"
            CisDB.SQL = gSL_Select
            If Not CisDB.DBRead Then
                Call KTTInsert
            End If
        Next
    End With
    
    KTTInsert_Proc = True
End Function
Public Function KTTInsertMnt_Proc() As Boolean
    KTTInsertMnt_Proc = False
    Dim wYosouYm    As String
    Dim wJiskiYm    As String
    
    gCnt2 = 0
    For gCnt1 = CisFun.Val2(iB1_TSYm) To CisFun.Val2(iB1_TEYm)
        If gCnt1 Mod 100 > 12 Then
            gCnt1 = gCnt1 + 88
        End If
        
        gCnt2 = gCnt2 + 1
        gSL_Select = " SELECT 予想年月,実績年月 FROM 型費テーブル"
        gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & "   AND 連番 = " & SaveRenban & ""
        gSL_Select = gSL_Select & "   AND 対象開始年月 = '" & Format(gCnt1, "000000") & "'"
        If Not KTTRead(gSL_Select, 1) Then
'*            If RTrim(iB1_KYm) = "" Or (RTrim(iB1_KYm) <> "" And Format(gCnt1, "000000") < RTrim(iB1_KYm)) Then
'*2005.04.07 Del
                Call ItemsClearKTT
                With KTT
                    .取引先 = RTrim(iH1_Torcd)
                    .品番 = RTrim(iH1_Hinbn)
                    .連番 = SaveRenban
                    .適用開始年月 = RTrim(iB1_TSYm)
                    .適用終了年月 = RTrim(iB1_TEYm)
                    .月発生金額 = CisFun.Val2(iB1_MmKin)
                    .投資金額 = CisFun.Val2(iB1_ToKin)
                    .回数 = RTrim(iB1_Kaisu)
                    .調整金額 = CisFun.Val2(B1lb_CKingk)
                    .切替年月 = RTrim(iB1_KYm)
                
                    .対象開始年月 = Format(gCnt1, "000000")
                    If gCnt2 > CisFun.Val2(iB1_Kaisu) Then
                        .金額 = CisFun.Val2(iB1_ToKin) - (CisFun.Val2(iB1_MmKin) * CisFun.Val2(iB1_Kaisu))
                    Else
                        .金額 = CisFun.Val2(iB1_MmKin)
                    End If
                    .作成者 = gTanto
                End With
                Call KTTInsert
'*            End If                         '2005.04.07 Del
        Else
            wJiskiYm = RTrim(KTT.実績年月)
            wYosouYm = RTrim(KTT.予想年月)
            If RTrim(iB1_KYm) <> "" Or Format(gCnt1, "000000") >= RTrim(iB1_KYm) Then
                If RTrim(KTT.実績年月) = "" Then
                    gSL_Select = " UPDATE 型費テーブル SET "
                    gSL_Select = gSL_Select & " 実績年月 = 'K'"
                    gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
                    gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
                    gSL_Select = gSL_Select & "   AND 連番 = " & SaveRenban & ""
                    gSL_Select = gSL_Select & "   AND 対象開始年月 = '" & Format(gCnt1, "000000") & "'"
                    CisDB.SQL = gSL_Select
                    CisDB.DBExec
                End If
            End If
            If RTrim(iB1_KYm) = "" Or Format(gCnt1, "000000") < RTrim(iB1_KYm) Then
                If RTrim(KTT.実績年月) = "K" Then
                    gSL_Select = " UPDATE 型費テーブル SET "
                    gSL_Select = gSL_Select & " 実績年月 = ''"
                    gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
                    gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
                    gSL_Select = gSL_Select & "   AND 連番 = " & SaveRenban & ""
                    gSL_Select = gSL_Select & "   AND 対象開始年月 = '" & Format(gCnt1, "000000") & "'"
                    CisDB.SQL = gSL_Select
                    CisDB.DBExec
                End If
            End If
            '予想済・実績済の赤の作成     （連番 + 100)
            gSL_Select = " SELECT 予想年月,実績年月 FROM 型費テーブル"
            gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
            gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
            gSL_Select = gSL_Select & "   AND 連番 = " & SaveRenban + 100 & ""
            gSL_Select = gSL_Select & "   AND 対象開始年月 = '" & Format(gCnt1, "000000") & "'"
            If Not KTTRead(gSL_Select, 1) Then
                If RTrim(iB1_KYm) <> "" And Format(gCnt1, "000000") >= RTrim(iB1_KYm) Then
                    Call ItemsClearKTT
                    With KTT
                        .取引先 = RTrim(iH1_Torcd)
                        .品番 = RTrim(iH1_Hinbn)
                        .連番 = SaveRenban + 100
                        .適用開始年月 = RTrim(iB1_TSYm)
                        .適用終了年月 = RTrim(iB1_TEYm)
                        .月発生金額 = CisFun.Val2(iB1_MmKin) * -1
                        .投資金額 = CisFun.Val2(iB1_ToKin) * -1
                        .回数 = RTrim(iB1_Kaisu)
                        .調整金額 = CisFun.Val2(B1lb_CKingk)
                        .切替年月 = RTrim(iB1_KYm)
                        If RTrim(wJiskiYm) = "" Then
                            .実績年月 = "K"
                        End If
                    
                        .対象開始年月 = Format(gCnt1, "000000")
                        If gCnt2 > CisFun.Val2(iB1_Kaisu) Then
                            .金額 = (CisFun.Val2(iB1_ToKin) - (CisFun.Val2(iB1_MmKin) * CisFun.Val2(iB1_Kaisu))) * -1
                        Else
                            .金額 = CisFun.Val2(iB1_MmKin) * -1
                        End If
                        .作成者 = gTanto
                    End With
                    Call KTTInsert
                End If
            Else
                If RTrim(wJiskiYm) = "" And (RTrim(iB1_KYm) <> "" Or Format(gCnt1, "000000") >= RTrim(iB1_KYm)) Then
                    If RTrim(KTT.実績年月) = "" Then
                        gSL_Select = " UPDATE 型費テーブル SET "
                        gSL_Select = gSL_Select & " 実績年月 = 'K'"
                        gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
                        gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
                        gSL_Select = gSL_Select & "   AND 連番 = " & SaveRenban + 100 & ""
                        gSL_Select = gSL_Select & "   AND 対象開始年月 = '" & Format(gCnt1, "000000") & "'"
                        CisDB.SQL = gSL_Select
                        CisDB.DBExec
                    End If
                End If
                If RTrim(iB1_KYm) = "" Or Format(gCnt1, "000000") < RTrim(iB1_KYm) Then
                    If RTrim(KTT.実績年月) = "K" Then
                        If RTrim(KTT.予想年月) = "" Then
                            gSL_Select = " DELETE FROM 型費テーブル "
                            gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
                            gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
                            gSL_Select = gSL_Select & "   AND 連番 = " & SaveRenban + 100 & ""
                            gSL_Select = gSL_Select & "   AND 対象開始年月 = '" & Format(gCnt1, "000000") & "'"
                            CisDB.SQL = gSL_Select
                            CisDB.DBExec
                        Else
                            gSL_Select = " UPDATE 型費テーブル SET "
                            gSL_Select = gSL_Select & " 実績年月 = ''"
                            gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
                            gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
                            gSL_Select = gSL_Select & "   AND 連番 = " & SaveRenban + 100 & ""
                            gSL_Select = gSL_Select & "   AND 対象開始年月 = '" & Format(gCnt1, "000000") & "'"
                            CisDB.SQL = gSL_Select
                            CisDB.DBExec
                        End If
                    End If
                End If
            End If
        End If
        
                
    Next
    
    KTTInsertMnt_Proc = True
End Function
Public Function KTTUpdate_Proc() As Boolean
    KTTUpdate_Proc = False
    With KTT
        gSL_Select = "Update 型費テーブル  Set "
        gSL_Select = gSL_Select & "切替年月 = '" & RTrim(iB1_KYm) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(gTanto) & "',"
        gSL_Select = gSL_Select & "更新端末 = HOST_NAME()"
        gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & "   AND 連番 IN (" & SaveRenban & "," & SaveRenban + 100 & ")"
    End With
    
    With CisDB
        .ConnectNo = 0
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KTTUpdate_Proc = True
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
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)

        Call .FixedSet(中中, 右中, あり, 3, 0, "No")
        Call .FixedSet(中中, 中中, あり, 12, 1, "適用開始年月")
        Call .FixedSet(中中, 中中, あり, 12, 2, "適用終了年月")
        Call .FixedSet(中中, 右中, あり, 14, 3, "月発生金額")
        Call .FixedSet(中中, 右中, あり, 5, 4, "回数")
        Call .FixedSet(中中, 右中, あり, 14, 5, "調整金額")
        Call .FixedSet(中中, 右中, あり, 15, 6, "投資金額")
        Call .FixedSet(中中, 中中, あり, 12, 7, "切替年月")
        
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 中中, あり, 0, 8, "X連番")
        Call .FixedSet(中中, 右中, あり, 0, 9, "X適用開始年月")
        Call .FixedSet(中中, 右中, あり, 0, 10, "X適用終了年月")
        Call .FixedSet(中中, 中中, あり, 0, 11, "X切替年月")
        Call .FixedSet(中中, 左中, あり, 0, 12, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp

    End With

End Sub
'****************************************************************************************'
'*      グリッド内容セット
'****************************************************************************************'
Private Sub Grid1Set()
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
    lb_Kensu = Format(Ttl_Kensu, "#,###")
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGrid1
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid1.FixedGet(ID)
        Next ID
        If Ttl_Kensu + 1 < CisVsGrid1.DispMax Then
            .Rows = CisVsGrid1.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    
    ID = 0
    
    Do Until Not KTT_RDSTS
        With VsGrid1
            ID = ID + 1
            '№
            .TextMatrix(ID, GridCol(0)) = Format(ID, "##")
            ' 適用開始年月
            If Trim(KTT.適用開始年月) <> "" Then
               .TextMatrix(ID, GridCol(1)) = Left(KTT.適用開始年月, 4) & "/" & Mid(KTT.適用開始年月, 5, 2)
            End If
            ' 適用終了年月
            If Trim(KTT.適用終了年月) <> "" Then
               .TextMatrix(ID, GridCol(2)) = Left(KTT.適用終了年月, 4) & "/" & Mid(KTT.適用終了年月, 5, 2)
            End If
            '月発生金額
            If KTT.月発生金額 <> 0 Then
                .TextMatrix(ID, GridCol(3)) = Format(KTT.月発生金額, "#,###")
            End If
            '回数
            If KTT.回数 <> 0 Then
                .TextMatrix(ID, GridCol(4)) = Format(KTT.回数, "##")
            End If
            '最終月調整金額
            If KTT.調整金額 <> 0 Then
                .TextMatrix(ID, GridCol(5)) = Format(KTT.調整金額, "#,###")
            End If
            '月発生金額
            If KTT.投資金額 <> 0 Then
                .TextMatrix(ID, GridCol(6)) = Format(KTT.投資金額, "#,###")
            End If
            ' 切替年月
            If Trim(KTT.切替年月) <> "" Then
               .TextMatrix(ID, GridCol(7)) = Left(KTT.切替年月, 4) & "/" & Mid(KTT.切替年月, 5, 2)
            End If
            
            '連番
            .TextMatrix(ID, GridCol(8)) = KTT.連番
            '適用開始年月
            .TextMatrix(ID, GridCol(9)) = RTrim(KTT.適用開始年月)
            '適用終了年月
            .TextMatrix(ID, GridCol(10)) = RTrim(KTT.適用終了年月)
            '切替年月
            .TextMatrix(ID, GridCol(11)) = RTrim(KTT.切替年月)
            
        End With
        
        Call KTTReadNext
    Loop
    
    Call KTTClose
    
    With VsGrid1
        .Row = 1: .Col = 1: .ColSel = .Cols - 1
        .Redraw = True
    End With
End Sub
'+---------------------------------------------------------------------------------------+
'+  グリッドより入力域へセット
'+---------------------------------------------------------------------------------------+
Private Sub GridToText()
    
    With VsGrid1
        If Trim(.TextMatrix(.Row, CisVsGrid1.FixedGet(8))) = "" Then Exit Sub
        
        iB1_TSYm = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(9)))
        iB1_TEYm = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(10)))
        iB1_MmKin = Format(.TextMatrix(.Row, CisVsGrid1.FixedGet(3)), "###,###,###")
        iB1_Kaisu = Format(.TextMatrix(.Row, CisVsGrid1.FixedGet(4)), "##")
        B1lb_CKingk = Format(.TextMatrix(.Row, CisVsGrid1.FixedGet(5)), "###,###,###")
        iB1_ToKin = Format(.TextMatrix(.Row, CisVsGrid1.FixedGet(6)), "#,###,###,###")
        iB1_KYm = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(11)))
    End With

End Sub
'+---------------------------------------+
'+  グリッド (Click)
'+---------------------------------------+
Private Sub VSGrid1_Click()
    With VsGrid1
        If ProcHB = "B2" Then
            Call HeadBodyClear("B1")

            Call GridToText
        End If
    End With
End Sub
'+---------------------------------------+
'+  グリッド (DblClick)
'+---------------------------------------+
Private Sub vsGrid1_DblClick()
    
    With VsGrid1
        If ProcHB = "B2" Then
            If Trim(.TextMatrix(.Row, CisVsGrid1.FixedGet(13))) <> "" Then
               Call ReturnPress
            End If
        End If
    End With
End Sub
'+---------------------------------------+
'+  グリッド (GotFocus)
'+---------------------------------------+
Private Sub vsGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace And Trim(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid1.FixedGet(13))) <> "" Then
       Call ReturnPress
    End If
End Sub
'+---------------------------------------+
'+  グリッド (RowColChange)
'+---------------------------------------+
Private Sub vsGrid1_RowColChange()
    If Not FormAct Then Exit Sub
    If ProcHB <> "B2" Then Exit Sub
    If RowColFlg Then Exit Sub
    Call VSGrid1_Click
End Sub
'+---------------------------------------+
'+  グリッド 再表示
'+---------------------------------------+
Private Sub vsGrid1_ReSetDisp()
    
    SaveRow = VsGrid1.Row
    RowColFlg = True
    gSL_Select = SaveSelect
    Call DBInput("INV")
    RowColFlg = False
    If SaveRow > VsGrid1.Rows - 1 Then
        SaveRow = SaveRow - 1
    End If
    If SyoriKB <> "ADD" Then
        VsGrid1.Row = SaveRow
    End If
    If VsGrid1.RowIsVisible(SaveRow) = False Then
        VsGrid1.TopRow = SaveRow
    End If

End Sub




