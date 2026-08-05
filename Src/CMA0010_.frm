VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{1BA4D453-0150-11CE-89B0-0000C037528B}#3.2#0"; "PDQCom32.OCX"
Begin VB.Form CMA0010 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "受注納品書入力 [丸高] "
   ClientHeight    =   11235
   ClientLeft      =   2865
   ClientTop       =   1455
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   525
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
      Left            =   510
      TabIndex        =   66
      Text            =   "X"
      Top             =   570
      Width           =   345
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
      Left            =   8595
      ScaleHeight     =   660
      ScaleWidth      =   5190
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   480
      Width           =   5250
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   2610
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
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
         Left            =   3465
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
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
         Left            =   4320
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
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
         Left            =   1755
         Top             =   30
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
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
         Left            =   885
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnCaption      =   "手入力"
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
      Begin CisBtn_60.CisBtn PB_RED 
         Height          =   585
         Left            =   30
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnCaption      =   "読取"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   8388736
         BtnShadowColor  =   16777215
         PFCaption       =   "( F5 )"
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
         PFBackColor     =   8388736
      End
   End
   Begin VB.PictureBox B2_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   555
      Left            =   4980
      ScaleHeight     =   495
      ScaleWidth      =   975
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1035
      Begin VB.TextBox iB2_Input 
         Height          =   330
         Left            =   60
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   75
         Width           =   825
      End
   End
   Begin VB.PictureBox B1_Area3 
      Height          =   1005
      Left            =   4575
      ScaleHeight     =   945
      ScaleWidth      =   1380
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1440
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   840
         Index           =   8
         Left            =   60
         Top             =   60
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1482
         ForeColor       =   16711680
         Caption         =   "ページ"
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
         Begin CisText_V60.CisText iB1_PageS 
            Height          =   375
            Left            =   135
            TabIndex        =   60
            Top             =   360
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   661
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
            Text            =   "99"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFbComma        =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
         Begin CisText_V60.CisText iB1_PageE 
            Height          =   375
            Left            =   750
            TabIndex        =   61
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
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
            Text            =   "99"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFbComma        =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "/"
            Height          =   315
            Left            =   570
            TabIndex        =   62
            Top             =   465
            Width           =   105
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   630
      ScaleHeight     =   1890
      ScaleWidth      =   3675
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   3435
      Width           =   3735
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   1725
         Left            =   90
         Top             =   90
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   3043
         ForeColor       =   16711680
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
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
         cAlingnment     =   6
         cBoderStyle     =   1
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   840
            Index           =   3
            Left            =   1170
            Top             =   30
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   1482
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
            cPositionX      =   100
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   345
               Left            =   1185
               Top             =   15
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   609
               ForeColor       =   16711680
               Caption         =   "変換先ｺｰﾄﾞ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ Ｐゴシック"
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
               cAlingnment     =   1
               cBoderStyle     =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D3 
               Height          =   300
               Left            =   855
               Top             =   405
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   529
               Caption         =   "→"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   15.75
                  Charset         =   128
                  Weight          =   700
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
            Begin Cis3D_v60.CIS3D CIS3D1 
               Height          =   375
               Left            =   1260
               Top             =   360
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
               BackColor       =   16777152
               Caption         =   "XXXXXXX"
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
               cPositionX      =   30
            End
            Begin CisText_V60.CisText iB1_NTorcd 
               Height          =   360
               Left            =   90
               TabIndex        =   54
               Top             =   360
               Width           =   735
               _ExtentX        =   1296
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
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   840
            Index           =   2
            Left            =   2430
            Top             =   870
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   1482
            ForeColor       =   16711680
            Caption         =   "受注者"
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
            Begin CisText_V60.CisText iB1_JTorcd 
               Height          =   375
               Left            =   165
               TabIndex        =   55
               Top             =   375
               Width           =   720
               _ExtentX        =   1270
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
               Text            =   "XXXXX"
               MaxLength       =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   840
            Index           =   0
            Left            =   30
            Top             =   30
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   1482
            ForeColor       =   16711680
            Caption         =   "納入番号"
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
            Begin CisText_V60.CisText iB1_Noubn 
               Height          =   375
               Left            =   195
               TabIndex        =   56
               Top             =   375
               Width           =   750
               _ExtentX        =   1323
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
               Text            =   "XXXXX"
               MaxLength       =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   840
            Index           =   12
            Left            =   30
            Top             =   870
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   1482
            ForeColor       =   16711680
            Caption         =   "納入指示日-便"
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
            Begin CisYMD_With_Btn.CisYMDwB iB1_NSYmd 
               Height          =   360
               Left            =   75
               TabIndex        =   57
               Top             =   360
               Width           =   1845
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
               Object.Width           =   1845
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
            Begin CisText_V60.CisText iB1_NBin 
               Height          =   375
               Left            =   1935
               TabIndex        =   58
               Top             =   360
               Width           =   345
               _ExtentX        =   609
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
               Text            =   "99"
               MaxLength       =   2
               cDataType       =   1
               cDataReplace    =   1
               cFaZero         =   0
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "##"
               cILength        =   2
            End
         End
      End
   End
   Begin VB.PictureBox B1_Area2 
      BackColor       =   &H00C0C0C0&
      Height          =   8895
      Left            =   6390
      ScaleHeight     =   8835
      ScaleWidth      =   7410
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1410
      Width           =   7470
      Begin VB.PictureBox B1_Area9 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8745
         Left            =   75
         ScaleHeight     =   8685
         ScaleWidth      =   7215
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   45
         Width           =   7275
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   360
            Left            =   0
            Top             =   0
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "品     番"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D UC_3D23 
            Height          =   360
            Left            =   3840
            Top             =   0
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "収容数"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D UC_3D21 
            Height          =   360
            Left            =   4965
            Top             =   0
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "箱数"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D UC_3D18 
            Height          =   360
            Left            =   2955
            Top             =   0
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "背番号"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D UC_3D26 
            Height          =   360
            Left            =   5850
            Top             =   0
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   635
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "納入数"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   0
            Left            =   4965
            Top             =   360
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   0
               Left            =   210
               TabIndex        =   3
               Top             =   105
               Width           =   495
               _ExtentX        =   873
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
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   0
            Left            =   3840
            Top             =   360
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   0
               Left            =   225
               TabIndex        =   4
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
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
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   0
            Left            =   2955
            Top             =   360
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   0
               Left            =   120
               TabIndex        =   5
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   0
            Left            =   0
            Top             =   360
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   0
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "１"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   0
               Left            =   525
               TabIndex        =   6
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XXXXXXX8"
               MaxLength       =   18
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   0
            Left            =   5850
            Top             =   360
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   0
               Left            =   210
               TabIndex        =   7
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
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
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   1
            Left            =   4965
            Top             =   915
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   1
               Left            =   210
               TabIndex        =   8
               Top             =   105
               Width           =   495
               _ExtentX        =   873
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
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   1
            Left            =   3840
            Top             =   915
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   1
               Left            =   225
               TabIndex        =   9
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
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
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   1
            Left            =   2955
            Top             =   915
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   1
               Left            =   120
               TabIndex        =   10
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   1
            Left            =   0
            Top             =   915
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   1
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "２"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   1
               Left            =   525
               TabIndex        =   11
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XXXXXXX8"
               MaxLength       =   18
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   1
            Left            =   5850
            Top             =   915
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   1
               Left            =   210
               TabIndex        =   12
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
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
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   2
            Left            =   4965
            Top             =   1470
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   2
               Left            =   210
               TabIndex        =   13
               Top             =   105
               Width           =   495
               _ExtentX        =   873
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
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   2
            Left            =   3840
            Top             =   1470
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   2
               Left            =   225
               TabIndex        =   14
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
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
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   2
            Left            =   2955
            Top             =   1470
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   2
               Left            =   120
               TabIndex        =   15
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   2
            Left            =   0
            Top             =   1470
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   2
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "３"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   2
               Left            =   525
               TabIndex        =   16
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XXXXXXX8"
               MaxLength       =   18
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   2
            Left            =   5850
            Top             =   1470
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   2
               Left            =   210
               TabIndex        =   17
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
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
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   3
            Left            =   4965
            Top             =   2025
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   3
               Left            =   210
               TabIndex        =   18
               Top             =   105
               Width           =   495
               _ExtentX        =   873
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
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   3
            Left            =   3840
            Top             =   2025
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   3
               Left            =   225
               TabIndex        =   19
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
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
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   3
            Left            =   2955
            Top             =   2025
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   3
               Left            =   120
               TabIndex        =   20
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   3
            Left            =   0
            Top             =   2025
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   3
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "４"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   3
               Left            =   525
               TabIndex        =   21
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XXXXXXX8"
               MaxLength       =   18
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   3
            Left            =   5850
            Top             =   2025
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   3
               Left            =   210
               TabIndex        =   22
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
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
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   4
            Left            =   4965
            Top             =   2580
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   4
               Left            =   210
               TabIndex        =   23
               Top             =   105
               Width           =   495
               _ExtentX        =   873
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
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   4
            Left            =   3840
            Top             =   2580
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   4
               Left            =   225
               TabIndex        =   24
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
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
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   4
            Left            =   2955
            Top             =   2580
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   4
               Left            =   120
               TabIndex        =   25
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   4
            Left            =   0
            Top             =   2580
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   4
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "５"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   4
               Left            =   525
               TabIndex        =   26
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XXXXXXX8"
               MaxLength       =   18
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   4
            Left            =   5850
            Top             =   2580
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   4
               Left            =   210
               TabIndex        =   27
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
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
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   5
            Left            =   4965
            Top             =   3135
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   5
               Left            =   210
               TabIndex        =   28
               Top             =   105
               Width           =   495
               _ExtentX        =   873
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
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   5
            Left            =   3840
            Top             =   3135
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   5
               Left            =   225
               TabIndex        =   29
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
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
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   5
            Left            =   2955
            Top             =   3135
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   5
               Left            =   120
               TabIndex        =   30
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   5
            Left            =   0
            Top             =   3135
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   5
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "６"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   5
               Left            =   525
               TabIndex        =   31
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XXXXXXX8"
               MaxLength       =   18
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   5
            Left            =   5850
            Top             =   3135
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   5
               Left            =   210
               TabIndex        =   32
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
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
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   6
            Left            =   4965
            Top             =   3690
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   6
               Left            =   210
               TabIndex        =   33
               Top             =   105
               Width           =   495
               _ExtentX        =   873
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
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   6
            Left            =   3840
            Top             =   3690
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   6
               Left            =   225
               TabIndex        =   34
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
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
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   6
            Left            =   2955
            Top             =   3690
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   6
               Left            =   120
               TabIndex        =   35
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   6
            Left            =   0
            Top             =   3690
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   6
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "７"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   6
               Left            =   525
               TabIndex        =   36
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XXXXXXX8"
               MaxLength       =   18
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   6
            Left            =   5850
            Top             =   3690
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   6
               Left            =   210
               TabIndex        =   37
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
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
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   7
            Left            =   4965
            Top             =   4245
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   7
               Left            =   210
               TabIndex        =   38
               Top             =   105
               Width           =   495
               _ExtentX        =   873
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
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   7
            Left            =   3840
            Top             =   4245
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   7
               Left            =   225
               TabIndex        =   39
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
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
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   7
            Left            =   2955
            Top             =   4245
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   7
               Left            =   120
               TabIndex        =   40
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   7
            Left            =   0
            Top             =   4245
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   7
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "８"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   7
               Left            =   525
               TabIndex        =   41
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XXXXXXX8"
               MaxLength       =   18
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   7
            Left            =   5850
            Top             =   4245
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   7
               Left            =   210
               TabIndex        =   42
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
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
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   8
            Left            =   4965
            Top             =   4800
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   8
               Left            =   210
               TabIndex        =   43
               Top             =   105
               Width           =   495
               _ExtentX        =   873
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
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   8
            Left            =   3840
            Top             =   4800
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   8
               Left            =   225
               TabIndex        =   44
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
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
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   8
            Left            =   2955
            Top             =   4800
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   8
               Left            =   120
               TabIndex        =   45
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   8
            Left            =   0
            Top             =   4800
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   8
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "９"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   8
               Left            =   525
               TabIndex        =   46
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XXXXXXX8"
               MaxLength       =   18
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   8
            Left            =   5850
            Top             =   4800
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   8
               Left            =   210
               TabIndex        =   47
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
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
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   9
            Left            =   4965
            Top             =   5355
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   9
               Left            =   210
               TabIndex        =   48
               Top             =   105
               Width           =   495
               _ExtentX        =   873
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
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   9
            Left            =   3840
            Top             =   5355
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   9
               Left            =   225
               TabIndex        =   49
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
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
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   9
            Left            =   2955
            Top             =   5355
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   9
               Left            =   120
               TabIndex        =   50
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   9
            Left            =   0
            Top             =   5355
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   9
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "10"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   9
               Left            =   525
               TabIndex        =   51
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XXXXXXX8"
               MaxLength       =   18
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   9
            Left            =   5850
            Top             =   5355
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   9
               Left            =   210
               TabIndex        =   52
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
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
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   10
            Left            =   4965
            Top             =   5910
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   10
               Left            =   210
               TabIndex        =   78
               Top             =   105
               Width           =   495
               _ExtentX        =   873
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
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   10
            Left            =   3840
            Top             =   5910
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   10
               Left            =   225
               TabIndex        =   79
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
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
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   10
            Left            =   2955
            Top             =   5910
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   10
               Left            =   120
               TabIndex        =   80
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   10
            Left            =   0
            Top             =   5910
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   10
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "11"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   10
               Left            =   525
               TabIndex        =   81
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XXXXXXX8"
               MaxLength       =   18
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   10
            Left            =   5850
            Top             =   5910
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   10
               Left            =   210
               TabIndex        =   82
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
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
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   11
            Left            =   4965
            Top             =   6465
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   11
               Left            =   210
               TabIndex        =   83
               Top             =   105
               Width           =   495
               _ExtentX        =   873
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
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   11
            Left            =   3840
            Top             =   6465
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   11
               Left            =   225
               TabIndex        =   84
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
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
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   11
            Left            =   2955
            Top             =   6465
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   11
               Left            =   120
               TabIndex        =   85
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   11
            Left            =   0
            Top             =   6465
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   11
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "12"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   11
               Left            =   525
               TabIndex        =   86
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XXXXXXX8"
               MaxLength       =   18
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   11
            Left            =   5850
            Top             =   6465
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   11
               Left            =   210
               TabIndex        =   87
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
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
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   12
            Left            =   4965
            Top             =   7020
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   12
               Left            =   210
               TabIndex        =   88
               Top             =   105
               Width           =   495
               _ExtentX        =   873
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
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   12
            Left            =   3840
            Top             =   7020
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   12
               Left            =   225
               TabIndex        =   89
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
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
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   12
            Left            =   2955
            Top             =   7020
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   12
               Left            =   120
               TabIndex        =   90
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   12
            Left            =   0
            Top             =   7020
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   12
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "13"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   12
               Left            =   525
               TabIndex        =   91
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XXXXXXX8"
               MaxLength       =   18
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   12
            Left            =   5850
            Top             =   7020
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   12
               Left            =   210
               TabIndex        =   92
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
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
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   13
            Left            =   4965
            Top             =   7575
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   13
               Left            =   210
               TabIndex        =   93
               Top             =   105
               Width           =   495
               _ExtentX        =   873
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
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   13
            Left            =   3840
            Top             =   7575
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   13
               Left            =   225
               TabIndex        =   94
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
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
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   13
            Left            =   2955
            Top             =   7575
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   13
               Left            =   120
               TabIndex        =   95
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   13
            Left            =   0
            Top             =   7575
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   13
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "14"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   13
               Left            =   525
               TabIndex        =   96
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XXXXXXX8"
               MaxLength       =   18
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   13
            Left            =   5850
            Top             =   7575
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   13
               Left            =   210
               TabIndex        =   97
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
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
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D22 
            Height          =   555
            Index           =   14
            Left            =   4965
            Top             =   8130
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   360
               Index           =   14
               Left            =   210
               TabIndex        =   98
               Top             =   105
               Width           =   495
               _ExtentX        =   873
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
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D24 
            Height          =   555
            Index           =   14
            Left            =   3840
            Top             =   8130
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   360
               Index           =   14
               Left            =   225
               TabIndex        =   99
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
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
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D20 
            Height          =   555
            Index           =   14
            Left            =   2955
            Top             =   8130
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   360
               Index           =   14
               Left            =   120
               TabIndex        =   100
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
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
               Text            =   "XXXX"
               MaxLength       =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D16 
            Height          =   555
            Index           =   14
            Left            =   0
            Top             =   8130
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
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
            Begin Cis3D_v60.CIS3D UC_3D3 
               Height          =   540
               Index           =   14
               Left            =   15
               Top             =   0
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   953
               BackColor       =   14737632
               Caption         =   "15"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin CisText_V60.CisText iB12_Hinbn 
               Height          =   360
               Index           =   14
               Left            =   525
               TabIndex        =   101
               Top             =   105
               Width           =   2340
               _ExtentX        =   4128
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
               Text            =   "XXXXXXXXX1XXXXXXX8"
               MaxLength       =   18
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D29 
            Height          =   555
            Index           =   14
            Left            =   5850
            Top             =   8130
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   360
               Index           =   14
               Left            =   210
               TabIndex        =   102
               Top             =   105
               Width           =   975
               _ExtentX        =   1720
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
               Text            =   "9999999"
               MaxLength       =   7
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#######"
               cILength        =   7
            End
         End
      End
   End
   Begin VB.Timer Timer2 
      Left            =   0
      Top             =   945
   End
   Begin Cis3D_v60.CIS3D UC_3D4 
      Height          =   405
      Left            =   7980
      Top             =   0
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   714
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
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cBoderWidth     =   3
      cBoderStyle     =   1
      Begin Cis3D_v60.CIS3D H1lb_ADDCnt 
         Height          =   300
         Left            =   3030
         Top             =   60
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         ForeColor       =   12583104
         Caption         =   "ZZ9"
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
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D UC_3D6 
         Height          =   300
         Left            =   1845
         Top             =   60
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         ForeColor       =   12583104
         Caption         =   "手 入 力"
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
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D H1lb_OCRCnt 
         Height          =   300
         Left            =   1230
         Top             =   60
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         Caption         =   "ZZ9"
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
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D UC_3D5 
         Height          =   300
         Left            =   45
         Top             =   60
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         Caption         =   "QR読取"
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
         cAlingnment     =   7
      End
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   15
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 受注納品書入力 [丸高]  】"
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
         Left            =   11700
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
   Begin PdqcommLib.PDQComm PDQComm1 
      Height          =   480
      Left            =   4275
      TabIndex        =   0
      Top             =   975
      Width           =   480
      _Version        =   196610
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.26
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FastScroll      =   0   'False
      Object.Height          =   34
      Object.Width           =   84
   End
   Begin Cis3D_v60.CIS3D K_Sykbnm 
      Height          =   795
      Left            =   645
      Top             =   480
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   1402
      Caption         =   "読取"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
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
   End
   Begin Cis3D_v60.CIS3D H1_Area1 
      Height          =   1920
      Left            =   630
      Top             =   1425
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   3387
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
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   870
         Left            =   2610
         Top             =   960
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   1535
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
      End
      Begin Cis3D_v60.CIS3D CIS3D17 
         Height          =   870
         Left            =   90
         Top             =   90
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   1535
         ForeColor       =   16711680
         Caption         =   "受注区分"
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
         cPositionX      =   90
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "指示"
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
            Left            =   2100
            TabIndex        =   74
            Top             =   120
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "特配"
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
            Index           =   6
            Left            =   2100
            TabIndex        =   73
            Top             =   480
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "臨時"
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
            Index           =   5
            Left            =   1170
            TabIndex        =   72
            Top             =   480
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H008080FF&
            Caption         =   "通常"
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
            Left            =   1170
            TabIndex        =   71
            Top             =   120
            Value           =   -1  'True
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "号試"
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
            Left            =   3030
            TabIndex        =   70
            Top             =   120
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "その他"
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
            Index           =   8
            Left            =   3960
            TabIndex        =   69
            Top             =   480
            Width           =   1320
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "スポット"
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
            Index           =   4
            Left            =   3960
            TabIndex        =   68
            Top             =   120
            Width           =   1320
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "代替"
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
            Index           =   7
            Left            =   3030
            TabIndex        =   67
            Top             =   480
            Width           =   840
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   870
         Index           =   0
         Left            =   90
         Top             =   960
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   1535
         ForeColor       =   16711680
         Caption         =   "－ 便  "
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
         cPositionX      =   120
         cPositionY      =   75
         Begin CisYMD_With_Btn.CisYMDwB iH1_Ymd 
            Height          =   345
            Left            =   90
            TabIndex        =   75
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
         Begin Cis3D_v60.CIS3D Back_Ymd 
            Height          =   300
            Left            =   510
            Top             =   75
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
         End
         Begin CisText_V60.CisText iH1_Bin 
            Height          =   360
            Left            =   2055
            TabIndex        =   76
            Top             =   420
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
      End
   End
   Begin VB.Label ReadColor 
      BackColor       =   &H00FFC0C0&
      Caption         =   "読取中色"
      Height          =   255
      Left            =   5070
      TabIndex        =   77
      Top             =   690
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "CMA0010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  受注データ取込み／フタバ産業
'**       フォームID    :  CSA0050
'**       処理概要      :
'**
'**       作  成  日    :  2004/03/05
'**       変  更  日    :  2004/06/02 By CIS 履歴テーブル出力
'**       変  更  日    :  2004/06/25 By CIS 1.受領書・出荷時発行の対応
'**       　  　  　                         2.処理追加（自社伝票＆ノーかんばんエフ発行）
'**
'**       変  更  日    :  2004/07/06-08 By TOTEC キャンセルデータリスト出力対応
'**       変  更  日    :  2004/07/09    By TOTEC 受注日を受信データ上の発注日とする
'**       変　更  日    :  2004/07/06  　By TOTEC 重複データが無い場合にキャンセルのデータ（メール）が存在する場合に帳票に出力
'**       変  更  日    :  2004/07/22  　By TOTEC 品番マスタ,品番取引先マスタに存在しない場合に帳票に出力
'**       変  更  日    :  2004/08/02  　By TOTEC 終了後処理での出荷指示書、ノーかんばんエフの発行中止
'**       変  更  日    :  2004/08/17  　By TOTEC 品番マスタ,品番取引先マスタに存在しない場合に帳票に出力のバグ対応
'**       変  更  日    :  2004/08/31 By CIS 1.品番取引先マスタの索引時、SK区分（発注区分：３桁目）を追加
'**       　  　  　                         2.受入は品番取引先マスタよりセット
'**       変  更  日    :  2004/09/09 By CIS 納品書発行区分セット(品取の伝種にて、対象外『9』のセット)
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl

    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
'
    Dim mSaveMsg        As String       ' ﾒｯｾｰｼﾞ

    Dim fs
    Dim ts
    Dim sFileName       As String
    Dim sDataOfLine     As String
    Dim nCountAll       As Long
    Dim nCountGet       As Long
    Dim nCountErr       As Long
'
    Dim wEndChk         As Boolean
    Dim wPos            As Integer
    Dim wDay            As Integer
    Dim wFlg            As String

    Dim mFileCode       As String
    Dim mHeadYmd        As String
    Dim mSeqNo          As String
    
    Dim mRirekiNo       As Long
    Const c_RL          As Integer = 200 + 2
    
'--↓2004/07/06--エラーリスト出力用＜TOTEC＞-------------------------------------------------
    '帳票設定定数
    Private Const mStrCrfFileName           As String = "CSA0050.crf"       'CoReportFileName
    
    Private Const mStrCrfDocName            As String = "フタバ受注データ取込み／エラーリスト"  'ドキュメント名
    
    Private Const mStrCrfPvFrmName          As String = "フタバ受注データ取込み／エラーリスト"  'プリンタダイアログ題名（プレビュー用）
    
    Private Const mStrOutPutOrient          As String = "Yoko"              'プリンタ向き ("Tate","Yoko")
    Private Const mStrOutPutSize            As String = "A4"                'PaperSize　 (A4,A3 etc...)
    Private Const mStrOutPutZoom            As Long = 100                   '印刷時の拡大縮小設定（％）
    Private Const mLngPrintKbn              As Long = 0                     '帳票区分（１プロジェクトに複数種類の帳票出力が有る場合に使用）
    Private Const mLngPageMax               As Long = 30                    'ページ内最大行数
    
    Private Const mcUserError               As Long = 65000                 'ユーザ発生エラー番号
    
    '帳票設定変数
    Private mLngPageCount                   As Long         'ページ数
    '品番未登録エラー用
    Private Const mStrCrfFileNameErr        As String = "CSA0050Err.crf"
    Private Const mStrCrfDocNameErr         As String = "品番未登録リスト"
    Private Const mStrCrfPvFrmNameErr       As String = "品番未登録リスト"
    Private mLngPageCountErr                   As Long         'ページ数
'--↑2004/07/06--エラーリスト出力用＜TOTEC＞------------------------------------------------


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
    mSaveMsg = lb_msg
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
    Timer2.Interval = 1000   ' 時刻表示間隔
    Timer2.Enabled = True    ' 時刻表示開始

    With CisFrm
        Set .MSGObject = lb_msg
        Set .MSGTimer = Timer2
        .MSGStopCount = 10
        .MSGTYPE = Left_To_Right
        .MSGText = mSaveMsg
        .MSGInterval = 100
        .MSGRelate
    End With

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   #     ロゴ表示       #
'   #-------------------#
    CisFrm.BGLDisp BG_Logo, 12
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    Set fs = CreateObject("Scripting.FileSystemObject")
'   初期設定の読込み
    Call Text_IniGet(sFileName)
    lb_FileName.Caption = sFileName

    If InitProc Then        ' 初期処理（データ削除）
        wEndChk = True
    Else
        wEndChk = False
        Call PB_END_Click
    End If
    
    mFileCode = "フタバ受注"
    mHeadYmd = Format(Now(), "YYYYMMDD")
    mSeqNo = "001"
        
    gSL_Select = "SELECT * FROM TEXT管理テーブル "
    gSL_Select = gSL_Select & " Where FILECD = '" & mFileCode & "'"
    If Not TKTRead(gSL_Select, 1) Then
    ' フタバ受注データ
        Call ItemsClearTKT
        With TKT
            .FILECD = mFileCode
            .F名称 = "フタバ受注データ"
            .作成者 = "SIMPR-A"
        End With
        Call TKTInsert
    End If
    
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ｐａｉｎｔ イベント》                   +
'+----------------------------------------------------------+
Private Sub Form_Paint()
'   ( グラデーション )
    CisFrm.Gradation 0, 150, 255, , , 6
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》       +
'+----------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Key_Acc = False

    If Not FormAct Then Exit Function
    If UCase(IVENT) <> "DOWN" Then Exit Function

    Select Case Key_Code
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn:    If PB_ENT.Visible Then Call PB_ENT_Click    '【入力】
           Case Else:           Exit Function
    End Select

    Key_Acc = True
End Function
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    If PB_END.Visible = False Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "    受注データ取込み中ですが処理を終了します。            "
            .MB_MSG(4) = "     よろしいですか？            "
            .MB_Title = "終了確認"
            .MB_Button = OK_CAN
            If Not .MBOX Then Exit Sub
        End With
        Call TERPutSak("J2")
        Call TextRirekiPutX(mRirekiNo, "T", "J2")
        
        Call CisDB.DBTran(TransRollback, 0)   ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    End If
    
    If wEndChk Then Call EndProc
    
    Call CisDB.DBDISConnect
    Call CisFrm.UnLoadDisp
    End
End Sub
'==========================================================='
'   終了処理（入力データがあれば後方処理を行う）
'==========================================================='
Private Sub EndProc()

'   << 入力件数確認 >>
    gSL_Select = ""
    gSL_Select = gSL_Select & "select count(*) 件数"
    gSL_Select = gSL_Select & "  from 受注入力ワーク"
    gSL_Select = gSL_Select & " where 端末番号 = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and"
    gSL_Select = gSL_Select & " (     出荷指示書NO = 0"
    gSL_Select = gSL_Select & "    or 納入番号 = ''"
    gSL_Select = gSL_Select & "    or 累積区分 = 0"
    gSL_Select = gSL_Select & " )"
    If JNWRead(gSL_Select, 1) Then
        If JNW.件数 > 0 Then
        '   出荷指示NO採番
'*--------------------------------------------------<< 2004/06/25 Insert Start >>
            gCHT_DelKbn53 = 0
'*--------------------------------------------------<< 2004/06/25 Insert End >>
            CXJ0110.Show vbModal
            Unload CXJ0110
            Set CXJ0110 = Nothing
            DoEvents
        '↓--2004/08/02 出荷指示書の発行中止 By TOTEC ---------------------------------------
        ''   出荷指示書発行
        '    Call CXJ0120LPrint(0)
        '↑--2004/08/02 出荷指示書の発行中止 By TOTEC ---------------------------------------
'*--------------------------------------------------<< 2004/06/25 Insert Start >>
        '   自社伝票発行
            Call CXJ0140LPrint(0)
        '   自社伝票ﾌｧｲﾙ出力
            Call CXJ0150TFile(0)
        '↓--2004/08/02 ノーかんばんエフの発行中止 By TOTEC ----------------------------------
        ''   ノーかんばんエフ発行
        '    Call CXJ0130LPrint(0)
        '↑--2004/08/02 ノーかんばんエフの発行中止 By TOTEC ----------------------------------
'*--------------------------------------------------<< 2004/06/25 Insert End >>
        End If
    End If

End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call ReturnPress
    Dummy.Enabled = False
End Sub
'+-------------------------------+
'+      テキスト入力先の変更
'+-------------------------------+
Private Sub PB_Dlg_Click()
    On Error Resume Next
    CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly
    CommonDialog1.FileName = sFileName
    CommonDialog1.Filter = "テキスト ファイル (*.txt)|*.txt"
    ' 以下のｿｰｽｺｰﾄﾞは、<ｷｬﾝｾﾙ>ﾎﾞﾀﾝがｸﾘｯｸされたときに実行されます。
    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen
    If Err = 32755 Then
        '<ｷｬﾝｾﾙ>ﾎﾞﾀﾝが押されたときにｴﾗｰｺｰﾄﾞを返すようにする設定です。
        '必要なｿｰｽｺｰﾄﾞを入力してください。
    End If
    If Trim(CommonDialog1.FileName) <> "" Then
        sFileName = CommonDialog1.FileName
        lb_FileName.Caption = sFileName
    End If
    On Error GoTo 0
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   Call CisFrm.TimeSet("1")
End Sub
'*******************************
'*      メッセージ表示        *
'*******************************
Private Sub Timer2_Timer()
    CisFrm.MSGRelate
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

    H1Mode = False: H1Color = gPLostSel
    B1Mode = False
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1": H1Mode = True: H1Color = gPGotSel
       Case "B1": B1Mode = True
    End Select
    H1_Area1.Enabled = H1Mode: H1_Area1.BackColor = H1Color

    If B1Mode Then PB_END.Move 6555
    PB_ENT.Visible = H1Mode
    PB_END.Visible = H1Mode Or B1Mode

    Timer2.Enabled = H1Mode Or B1Mode

    picBar.Visible = False

' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'*********************************
'*          入力処理(Enter)      *
'*********************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "B2"
        Call DispChange(ProcHB)
        If DBPut Then
            With CisFrm
                .MSGTYPE = FSize_Changes
                .MSGText = lb_msg.Caption
            End With
            ProcHB = "B1"
        Else
            ProcHB = "H1"
        End If
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

    If fs.FileExists(sFileName) Then
        Set ts = fs.GetFile(sFileName)
        nCountAll = ts.Size / c_RL
    Else
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    受注データが存在しません！            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If

    Head1Chk = True
End Function

'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+      追加処理
'+----------------------+
Private Function DBPut()
    DBPut = False

    lb_msg.Caption = mSaveMsg
    DoEvents
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "    受注データの取込みを実行します。            "
        .MB_MSG(4) = "     よろしいですか？            "
        .MB_Title = "変換処理"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With

    Back_Msg.Visible = True

    On Error GoTo DBPut_Err
    
    ' 履歴情報出力
    mRirekiNo = 0
    Call TextRirekiPutX(mRirekiNo, "S", "")
    
    Call CisDB.DBTran(TransBegin)    ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

    lb_msg.Caption = "【  受  注  デ　ー　タ　読　込　み　中  】"
    DoEvents
    If Not DBInsert_JNW Then
       Call TERPutSak("J1")
       Call TextRirekiPutX(mRirekiNo, "T", "J1")
       GoTo DBPut_Err
    End If
    Call CisDB.DBTran(TransCommit)   ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    On Error GoTo 0

    lb_msg.Caption = "【  処  理  終  了  】"
    DoEvents
    
    Call TextRirekiPutX(mRirekiNo, "E", "")

'   初期設定の書込み
    Call Text_IniPut(sFileName)

    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function

'+--------------------------------+
'+      フタバ受注データの取込
'+--------------------------------+
Private Function DBInsert_JNW() As Boolean
    
'--↓2004/07/06--エラーリスト出力用＜TOTEC＞-------------------------------------------------
    Dim Print_Flg As Boolean
    Dim Print_FlgErr As Boolean
    
    Dim ObjCoForm               As CrForm       'CoReportオブジェクト
    Dim ObjCoFormErr               As CrForm       'CoReportオブジェクト
    Dim StrErrMsg               As String       'エラーメッセージ
    Dim LngOutRow           As Long                 '出力行カウンタ
    Dim LngOutRowErr           As Long                 '出力行カウンタ
    Dim BlnOpenDocument         As Boolean      'CoReport印刷フラグ
     
    Dim Dkubun                  As String       'データ区分
'--↑2004/07/06--エラーリスト出力用＜TOTEC＞-------------------------------------------------
     
    DBInsert_JNW = False
    
'--↓2004/07/06--エラーリスト出力用＜TOTEC＞-------------------------------------------------
    Print_Flg = False
    Print_FlgErr = False
    
    'プリントアウト可能にする
    If Output_ReportStr(Me.CrDraw, _
             "", _
             "Print", BlnOpenDocument) = True Then
            'CRフォームオープン
'*---------------------------------------------------<< 2004/08/31 >> Update Start By CIS
'           Set ObjCoForm = Me.CrDraw.OpenForm(App.path & "\" & mStrCrfFileName)  'キャンセルエラー用
'           Set ObjCoFormErr = Me.CrDraw.OpenForm(App.path & "\" & mStrCrfFileNameErr)    '品番が存在しないエラー出力用
            Set ObjCoForm = Me.CrDraw.OpenForm(gIniExe & mStrCrfFileName)   'キャンセルエラー用
            Set ObjCoFormErr = Me.CrDraw.OpenForm(gIniExe & mStrCrfFileNameErr)    '品番が存在しないエラー出力用
'*---------------------------------------------------<< 2004/08/31 >> Update End   By CIS
            LngOutRow = 0
            LngOutRowErr = 0
    End If
    'プリントアウト可能にする 2004/07/22 追加
    If Output_ReportStrErr(Me.CrDraw1, _
             "", _
             "Print", BlnOpenDocument) = True Then
            'CRフォームオープン
'*---------------------------------------------------<< 2004/08/31 >> Update Start By CIS
'           Set ObjCoFormErr = Me.CrDraw.OpenForm(App.path & "\" & mStrCrfFileNameErr)    '品番が存在しないエラー出力用
            Set ObjCoFormErr = Me.CrDraw.OpenForm(gIniExe & mStrCrfFileNameErr)    '品番が存在しないエラー出力用
'*---------------------------------------------------<< 2004/08/31 >> Update End   By CIS
            LngOutRowErr = 0
    End If
    
    
    
'--↑2004/07/06--エラーリスト出力用＜TOTEC＞-------------------------------------------------
    
    
'   // Open //
    Set ts = fs.OpenTextFile(sFileName, 1, 0)
    nCountGet = 0
    nCountErr = 0
    picBar.Visible = True
    With CisFrm
        Set .MeterPic = picBar
        .MeterMax = nCountAll
    End With

'   // Read //
    Do While ts.AtEndOfStream <> True
        sDataOfLine = ts.ReadLine
        nCountGet = nCountGet + 1

'--↓2004/07/06--エラーリスト出力用にコメント化＜TOTEC＞-------------------------------------------------
''''
'''        If DBSet_Chk Then                       'データ存在チェック
'''           Call DBSet_Sy                        'データ追加
'''           B1lb_CntPut.Caption = Format(CisFun.Val2(B1lb_CntPut.Caption) + 1, "#,###")
'''        Else
'''           B1lb_CntErr.Caption = Format(CisFun.Val2(B1lb_CntErr.Caption) + 1, "#,###")
'''        End If
'''        B1lb_CntGet.Caption = Format(CisFun.Val2(B1lb_CntGet.Caption) + 1, "#,###")
'''        DoEvents
''''
'--↑2004/07/06--エラーリスト出力用にコメント化＜TOTEC＞-------------------------------------------------
          
'--↓2004/07/06--エラーリスト出力用＜TOTEC＞------------------------------------------------------------
    '↓--2004/08/17 追加＜TOTEC＞-------------------------------------------------------------------------
    '   <><><> 名称マスタ参照 <><><>
        Call ItemsClearSYM
        With SYM
            .区分名称 = "フタバ納入先管理"
            .区分桁数 = 1
            .値 = "1"
            .値名称 = "納入先"
            .英数字1 = "100"
        End With
        Call SYMReadChk
        If RTrim(SYM.英数字1) = "" Then SYM.英数字1 = "100"
        Call ItemsClearJNW
    '↓--2004/08/17 追加＜TOTEC＞-------------------------------------------------------------------------
       'データ区分格納(A:"発注",C:"取消")
        Dkubun = RTrim(CisFun.Mid2(sDataOfLine, 1, 1))
        
        '品番存在チェック
        If Chk_Hin Then
            If DBSet_Chk Then                       'データ存在チェック
            '受注入力ワーク内にダブりがない場合 ***
                If Dkubun = "C" Then
                'データ区分が"C"の場合、
                    'キャンセルリスト出力<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
                        With JNW
                            .受注日 = "20" & RTrim(CisFun.Mid2(sDataOfLine, 114, 6))
                            .納入日 = "20" & RTrim(CisFun.Mid2(sDataOfLine, 66, 6))
                            .納入便 = RTrim(CisFun.Mid2(sDataOfLine, 72, 2))
                            .納入先 = RTrim(SYM.英数字1) & (CisFun.Mid2(sDataOfLine, 100, 1))
                            .受入 = RTrim(CisFun.Mid2(sDataOfLine, 28, 2))
                            .納入番号 = RTrim(CisFun.Mid2(sDataOfLine, 60, 5))
                            .納品書行 = CisFun.Val2(RTrim(CisFun.Mid2(sDataOfLine, 65, 1)))
                            .品番 = RTrim(CisFun.Mid2(sDataOfLine, 35, 20))
                            .背番号 = RTrim(CisFun.Mid2(sDataOfLine, 31, 4))
                            .枚数 = CisFun.Val2(RTrim(CisFun.Mid2(sDataOfLine, 87, 3)))
                            .収容数 = CisFun.Val2(RTrim(CisFun.Mid2(sDataOfLine, 55, 5)))
                            .納入数 = CisFun.Val2(RTrim(CisFun.Mid2(sDataOfLine, 90, 6)))
                        End With
                        '出力処理
                        If Not Output_Main(ObjCoForm, "Print", LngOutRow, StrErrMsg) Then
            '             If Not Output_Main() Then
                            Err.Raise mcUserError
                        End If
                        LngOutRow = LngOutRow + 1           '出力行++1
                        '改ページ処理
                            If LngOutRow >= mLngPageMax Then
                               LngOutRow = 0
                            End If
                        Print_Flg = True
                    '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
                        
                        'エラー件数UP
                        B1lb_CntErr.Caption = Format(CisFun.Val2(B1lb_CntErr.Caption) + 1, "#,###")
                        
                Else
                'データ区分が"C"以外("A")の場合、
                    'データ追加
                    Call DBSet_Sy
                    B1lb_CntPut.Caption = Format(CisFun.Val2(B1lb_CntPut.Caption) + 1, "#,###")
                End If
            Else
            '受注入力ワーク内にダブりがある場合 ***
                With JNW
                    .受注日 = "20" & RTrim(CisFun.Mid2(sDataOfLine, 114, 6))
                    .納入日 = "20" & RTrim(CisFun.Mid2(sDataOfLine, 66, 6))
                    .納入便 = RTrim(CisFun.Mid2(sDataOfLine, 72, 2))
                    .納入先 = RTrim(SYM.英数字1) & (CisFun.Mid2(sDataOfLine, 100, 1))
                    .受入 = RTrim(CisFun.Mid2(sDataOfLine, 28, 2))
                    .納入番号 = RTrim(CisFun.Mid2(sDataOfLine, 60, 5))
                    .納品書行 = CisFun.Val2(RTrim(CisFun.Mid2(sDataOfLine, 65, 1)))
                    .品番 = RTrim(CisFun.Mid2(sDataOfLine, 35, 20))
                    .背番号 = RTrim(CisFun.Mid2(sDataOfLine, 31, 4))
                    .枚数 = CisFun.Val2(RTrim(CisFun.Mid2(sDataOfLine, 87, 3)))
                    .収容数 = CisFun.Val2(RTrim(CisFun.Mid2(sDataOfLine, 55, 5)))
                    .納入数 = CisFun.Val2(RTrim(CisFun.Mid2(sDataOfLine, 90, 6)))
                End With
                '受注データへ未累積の場合
                'データ区分が"C"の場合⇒削除/"A"の場合⇒更新
                If JNW.累積区分 = 0 Then
                    If Dkubun = "C" Then
                    'データ区分が"C"の場合、
                        '該当データを削除する
                        Call DBSet_Del
                    Else
                    'データ区分が"C"以外("A")の場合、
                        '該当データを更新する
                        Call DBSet_Upd
                    End If
                End If
    '           B1lb_CntErr.Caption = Format(CisFun.Val2(B1lb_CntErr.Caption) + 1, "#,###")
                '正常処理件数UP
                B1lb_CntPut.Caption = Format(CisFun.Val2(B1lb_CntPut.Caption) + 1, "#,###")
            End If
        Else
            '存在しなかった場合リスト出力 2004/07/22 追加<><><><><><><><><><><><><><><><><><><><><><>
            '↓--2004/08/17 追加TOTEC------------------------------------------------
                With JNW
                    '品番
                    .品番 = RTrim(CisFun.Mid2(sDataOfLine, 35, 20))
                    '背番号
                    .背番号 = RTrim(CisFun.Mid2(sDataOfLine, 31, 4))
                    '受入
                    .受入 = RTrim(CisFun.Mid2(sDataOfLine, 28, 2))
                    '納入番号
                    .納入番号 = RTrim(CisFun.Mid2(sDataOfLine, 60, 5))
                    '納入書行
                    .納品書行 = CisFun.Val2(RTrim(CisFun.Mid2(sDataOfLine, 65, 1)))
                    '枚数
                    .枚数 = CisFun.Val2(RTrim(CisFun.Mid2(sDataOfLine, 87, 3)))
                    '収容数
                    .収容数 = CisFun.Val2(RTrim(CisFun.Mid2(sDataOfLine, 55, 5)))
                    '納入数
                    .納入数 = CisFun.Val2(RTrim(CisFun.Mid2(sDataOfLine, 90, 6)))
                    '納入日
                    .納入日 = "20" & RTrim(CisFun.Mid2(sDataOfLine, 66, 6))
                    '納入便
                    .納入便 = RTrim(CisFun.Mid2(sDataOfLine, 72, 2))
                    '納入先
                    .納入先 = RTrim(SYM.英数字1) & (CisFun.Mid2(sDataOfLine, 100, 1))
                    '受注日
                    .受注日 = "20" & RTrim(CisFun.Mid2(sDataOfLine, 114, 6))
                End With
            '↑--2004/08/17 追加-TOTEC-----------------------------------------------
            '出力処理
            If Not Output_MainErr(ObjCoFormErr, "Print", LngOutRowErr, StrErrMsg) Then
                Err.Raise mcUserError
            End If
            LngOutRowErr = LngOutRowErr + 1           '出力行++1
            '改ページ処理
                If LngOutRowErr >= mLngPageMax Then
                   LngOutRowErr = 0
                End If
            Print_FlgErr = True
            '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
            'エラー件数UP
            B1lb_CntErr.Caption = Format(CisFun.Val2(B1lb_CntErr.Caption) + 1, "#,###")
        End If
            
        B1lb_CntGet.Caption = Format(CisFun.Val2(B1lb_CntGet.Caption) + 1, "#,###")
        DoEvents
'--↑2004/07/06--エラーリスト出力用＜TOTEC＞---------------------------------------------------------------

        CisFrm.MeterUpdate nCountGet
    Loop

'--↓2004/07/06--エラーリスト出力用＜TOTEC＞---------------------------------------------------------------
    If Print_Flg Then
        If Set_Kaikou(ObjCoForm, "Print", LngOutRow, StrErrMsg) = False Then
        End If
    End If
    If Print_FlgErr Then
        If Set_KaikouErr(ObjCoFormErr, "Print", LngOutRowErr, StrErrMsg) = False Then
        End If
    End If
    'プリントアウト終了
        ObjCoForm.Close
        ObjCoFormErr.Close
        If Output_ReportEnd(Me.CrDraw, ObjCoForm, _
                 "", _
                 "Print", BlnOpenDocument) = True Then
        End If
        ' 2004/07/22 追加--------------------------------
        If Output_ReportEnd(Me.CrDraw1, ObjCoFormErr, _
                 "", _
                 "Print", BlnOpenDocument) = True Then
        End If
        '------------------------------------------------
'--↑2004/07/06--エラーリスト出力用＜TOTEC＞---------------------------------------------------------------
    
    picBar.Visible = False
'   // Close //
    ts.Close

    DBInsert_JNW = True
End Function


'==============================================================================*
' プロシージャ名　　：Output_Report
' 機能　　　　　　　：帳票出力、統括処理
' 引数　　　　　　　：iObjCrDraw        CoReportsオブジェクト
'                     iStrPrinterNM     プリンタ名称
'                     iStrOutType       出力タイプ（Print：印刷、View：プレビュー）
'                     P_PRTKBN          印刷区分（不使用）
' 戻り値　　　　　　：Output_Report
'                       TRUE :正常終了
'                       FALSE:エラー
' 作成日／作成者　　：
' 備考　　　　　　　：
'
'------------------------------------------------------------------------------*
Public Function Output_ReportStr(ByRef iObjCrDraw As CrDraw, _
                              ByVal iStrPrinterNM As String, _
                              ByVal iStrOutType As String, _
                              ByVal BlnOpenDocument As Boolean, _
                              Optional ByVal P_PRTKBN As Integer = 0) As Boolean

'    Dim BlnOpenDocument         As Boolean      'CoReport印刷フラグ
    Dim StrErrMsg               As String       'エラーメッセージ
    Dim ObjCoForm               As CrForm       'CoReportオブジェクト

On Error GoTo ERR_PROC:
    
    '変数初期化
    Output_ReportStr = False
    
    'ページ番号初期化
    mLngPageCount = 1
    mLngPageCountErr = 1
    
    'ドキュメントオープンフラグの初期化
    BlnOpenDocument = False
    
    '出力初期設定
    If Not Output_Setting(iObjCrDraw, iStrPrinterNM, iStrOutType, StrErrMsg) Then
        Err.Raise mcUserError
    End If
    
    'ドキュメントフラグのセット
    BlnOpenDocument = True

Exit_PROC:
    
    Set ObjCoForm = Nothing
    Output_ReportStr = True
        
    Exit Function

ERR_PROC:
End Function
''==============================================================================*
'' プロシージャ名　　：Output_ReportErr
'' 機能　　　　　　　：帳票出力、統括処理(品番未登録リスト用)
'' 引数　　　　　　　：iObjCrDraw        CoReportsオブジェクト
''                     iStrPrinterNM     プリンタ名称
''                     iStrOutType       出力タイプ（Print：印刷、View：プレビュー）
''                     P_PRTKBN          印刷区分（不使用）
'' 戻り値　　　　　　：Output_Report
''                       TRUE :正常終了
''                       FALSE:エラー
'' 作成日／作成者　　：2004/07/26
'' 備考　　　　　　　：
''
''------------------------------------------------------------------------------*
Public Function Output_ReportStrErr(ByRef iObjCrDraw As CrDraw, _
                              ByVal iStrPrinterNM As String, _
                              ByVal iStrOutType As String, _
                              ByVal BlnOpenDocument As Boolean, _
                              Optional ByVal P_PRTKBN As Integer = 0) As Boolean

'    Dim BlnOpenDocument         As Boolean      'CoReport印刷フラグ
    Dim StrErrMsg               As String       'エラーメッセージ
    Dim ObjCoForm               As CrForm       'CoReportオブジェクト

On Error GoTo ERR_PROC:

    '変数初期化
    Output_ReportStrErr = False

    'ページ番号初期化
'    mLngPageCount = 1
    mLngPageCountErr = 1

    'ドキュメントオープンフラグの初期化
    BlnOpenDocument = False

    '出力初期設定
    If Not Output_SettingErr(iObjCrDraw, iStrPrinterNM, iStrOutType, StrErrMsg) Then
        Err.Raise mcUserError
    End If

    'ドキュメントフラグのセット
    BlnOpenDocument = True

Exit_PROC:

    Set ObjCoForm = Nothing
    Output_ReportStrErr = True

    Exit Function

ERR_PROC:
    Output_ReportStrErr = False

    'エラーメッセージ出力
    If Err.Number = mcUserError Then
        MsgBox StrErrMsg, vbOKOnly + vbExclamation, App.Title & "(Output_Report)"       'ユーザエラー
    Else
        MsgBox ErrMsg_VB(), vbOKOnly + vbExclamation, App.Title & "(Output_Report)"     'システムエラー
    End If

    '印刷 JOB の Abort
    If BlnOpenDocument = True Then
        If iStrOutType = "Print" Then
            iObjCrDraw.AbortDocument
            iObjCrDraw.EndPrinter
        End If
        ObjCoForm.Close
    End If

    GoTo Exit_PROC:

End Function
'==============================================================================*
' プロシージャ名　　：Output_Report
' 機能　　　　　　　：帳票出力、統括処理
' 引数　　　　　　　：iObjCrDraw        CoReportsオブジェクト
'                     iStrPrinterNM     プリンタ名称
'                     iStrOutType       出力タイプ（Print：印刷、View：プレビュー）
'                     P_PRTKBN          印刷区分（不使用）
' 戻り値　　　　　　：Output_Report
'                       TRUE :正常終了
'                       FALSE:エラー
' 作成日／作成者　　：
' 備考　　　　　　　：
'
'------------------------------------------------------------------------------*
Public Function Output_ReportEnd(ByRef iObjCrDraw As CrDraw, ObjCoForm As CrForm, _
                              ByVal iStrPrinterNM As String, _
                              ByVal iStrOutType As String, _
                              Optional ByVal P_PRTKBN As Integer = 0) As Boolean

    Dim BlnOpenDocument         As Boolean      'CoReport印刷フラグ
    Dim StrErrMsg               As String       'エラーメッセージ

On Error GoTo ERR_PROC:
    
    'CRフォームをクローズします。
'    ObjCoForm.Close

    'ドキュメントフラグを降ろす
    BlnOpenDocument = False

    '出力終了
    If Not Output_End(iObjCrDraw, iStrOutType, StrErrMsg) Then
        Err.Raise mcUserError
    End If

Exit_PROC:
    
    Set ObjCoForm = Nothing
    Output_ReportEnd = True
        
    Exit Function

ERR_PROC:

    Output_ReportEnd = False
    
    'エラーメッセージ出力
    If Err.Number = mcUserError Then
        MsgBox StrErrMsg, vbOKOnly + vbExclamation, App.Title & "(Output_Report)"       'ユーザエラー
    Else
        MsgBox ErrMsg_VB(), vbOKOnly + vbExclamation, App.Title & "(Output_Report)"     'システムエラー
    End If
    
    '印刷 JOB の Abort
    If BlnOpenDocument = True Then
        If iStrOutType = "Print" Then
            iObjCrDraw.AbortDocument
            iObjCrDraw.EndPrinter
        End If
        ObjCoForm.Close
    End If
    
    GoTo Exit_PROC:
    
End Function
Function ErrMsg_VB() As String

    Dim ERRMSG As String

    ERRMSG = ""
    ERRMSG = ERRMSG & Err.Source & " でエラーが発生しました。" & Chr(13)
    ERRMSG = ERRMSG & Chr(13)
    ERRMSG = ERRMSG & "エラー番号 ： " & Str(Err.Number) & Chr(13)
    ERRMSG = ERRMSG & "エラー内容 ： " & Err.Description & Chr(13)
    ERRMSG = ERRMSG & Chr(13)
    ERRMSG = ERRMSG & Chr(13)
    ERRMSG = ERRMSG & "上記のエラーの為、プログラムを中断します。" & Chr(13)
    ErrMsg_VB = ERRMSG

End Function
'*-----------------------------------------------------------------------------*
'* 機能：帳票出力初期設定
'* 概要：
'* 引数：iObjCrDraw       　　　　  Coreportsオブジェクト
'* 　　　iStrPrinterNM       　　   プリンタ名称
'* 　　　iStrPType       　　　　   Print：印刷、View：プレビュ－表示
'* 　　　oStrErrMsg     　　　　    エラーメッセージ
'*
'* 戻値：Output_Setting         TRUE：正常、FALSE：エラー
'*
'*-----------------------------------------------------------------------------*'
Function Output_Setting(ByRef iObjCrDraw As CrDraw, _
                        ByRef iStrPrinterNM As String, _
                        ByVal iStrPType As String, _
                        ByRef oStrErrMsg As String) As Boolean


    Dim LngRetunCod         As Long     'メソッド戻り値
    Dim LngZoom             As Integer  '拡大縮小率

On Error GoTo ErrorHandle
    
    Output_Setting = False
    
    '初期化処理
    LngRetunCod = iObjCrDraw.Initialize
    If LngRetunCod <> 0 Then
        Output_Setting = False
        oStrErrMsg = "プリンタ初期化エラー(" & sGetErrMessage(LngRetunCod) & ")"
        Exit Function
    End If
    
    '用紙向き設定
    Select Case mStrOutPutOrient
    Case "Tate"
        iObjCrDraw.Printer.Orientation = corOrientVertical
    Case Else
        iObjCrDraw.Printer.Orientation = corOrientHorizontal
    End Select
    
    '用紙サイズ設定
    Select Case mStrOutPutSize
    Case "A3": iObjCrDraw.Printer.PaperSize = corPaperA3
    Case "B4": iObjCrDraw.Printer.PaperSize = corPaperB4
    Case "A4": iObjCrDraw.Printer.PaperSize = corPaperA4
    Case "B5": iObjCrDraw.Printer.PaperSize = corPaperB5
    Case "A5": iObjCrDraw.Printer.PaperSize = corPaperA5
    Case Else: iObjCrDraw.Printer.PaperSize = corPaperA4  '上記以外は A4 に設定
    End Select
    
    Select Case iStrPType
    Case "Print"
        'プリンタ出力開始処理
        If iObjCrDraw.CrPrinters.Count = 0 Then
            Output_Setting = False
            oStrErrMsg = "プリンタが未設定です"
            Exit Function
        End If
        
        '拡大縮小率を設定
        iObjCrDraw.Printer.PrintScale = mStrOutPutZoom
        
        'プリンタ出力を設定
        LngRetunCod = iObjCrDraw.StartPrinter(iStrPrinterNM, mStrCrfDocName)
        
        '用紙向き設定
        Select Case mStrOutPutOrient
          Case "Tate"
            iObjCrDraw.Printer.Orientation = corOrientVertical
          Case Else
            iObjCrDraw.Printer.Orientation = corOrientHorizontal
         End Select
    
        '用紙サイズ設定
        Select Case mStrOutPutSize
         Case "A3": iObjCrDraw.Printer.PaperSize = corPaperA3
         Case "B4": iObjCrDraw.Printer.PaperSize = corPaperB4
         Case "A4": iObjCrDraw.Printer.PaperSize = corPaperA4
         Case "B5": iObjCrDraw.Printer.PaperSize = corPaperB5
         Case "A5": iObjCrDraw.Printer.PaperSize = corPaperA5
         Case Else: iObjCrDraw.Printer.PaperSize = corPaperA4  '上記以外は A4 に設定
        End Select
    
    
    Case "View"
        'プレビュー出力開始処理
        '拡大縮小率を設定
'        LngZoom = Val(PrinterDialog.txt_Zoom.Text)
        LngZoom = mStrOutPutZoom              'ダイアログが無い為、印刷の設定を使用
        iObjCrDraw.Printer.PrintScale = LngZoom

        '印刷用紙の大きさと同じ設定にする
        iObjCrDraw.Unit = corUnitTwip
        
        '以下保留（プレビューは使用しない？）
        'プレビューフォームサイズ設定
'        glFormWidth = mObjCoForm.PaperWidth * LngZoom / 100
'        glFormHeight = mObjCoForm.PaperLength * LngZoom / 100

        'プレビュー出力先領域設定(フォームの縦横でLとWを変える)
'        Set PreviewForm.Picture2(mLngPageCount) = Nothing
'        If mStrOutPutOrient = "Tate" Then
'            PreviewForm.Picture2(mLngPageCount).Width = mObjCoForm.PaperWidth * LngZoom / 100
'            PreviewForm.Picture2(mLngPageCount).Height = mObjCoForm.PaperLength * LngZoom / 100
'        ElseIf mStrOutPutOrient = "Yoko" Then
'            PreviewForm.Picture2(mLngPageCount).Width = mObjCoForm.PaperLength * LngZoom / 100
'            PreviewForm.Picture2(mLngPageCount).Height = mObjCoForm.PaperWidth * LngZoom / 100
'        End If

        DoEvents

        'デバイスコンテキスト出力設定処理
        LngRetunCod = iObjCrDraw.SetOutputDC(PreviewForm.Picture2(mLngPageCount).hDC)
    
    End Select

    'エラー判定
    If LngRetunCod <> 0 Then
        Output_Setting = False
        oStrErrMsg = "プリンタ出力先設定エラー(" & sGetErrMessage(LngRetunCod) & ")"
        Exit Function
    End If

    '座標系単位指定
    iObjCrDraw.Unit = corUnitMm
    
    Output_Setting = True
    
    Exit Function

ErrorHandle:
    
    'ｴﾗｰ ﾒｯｾｰｼﾞ取得
    oStrErrMsg = ErrMsg_VB
    
End Function

''*-----------------------------------------------------------------------------*
''* 機能：帳票出力初期設定(品番未登録リスト用)
''* 概要：
''* 引数：iObjCrDraw       　　　　  Coreportsオブジェクト
''* 　　　iStrPrinterNM       　　   プリンタ名称
''* 　　　iStrPType       　　　　   Print：印刷、View：プレビュ－表示
''* 　　　oStrErrMsg     　　　　    エラーメッセージ
''*
''* 戻値：Output_Setting         TRUE：正常、FALSE：エラー
''*作成日：2004/07/26
''*-----------------------------------------------------------------------------*'
Function Output_SettingErr(ByRef iObjCrDraw As CrDraw, _
                        ByRef iStrPrinterNM As String, _
                        ByVal iStrPType As String, _
                        ByRef oStrErrMsg As String) As Boolean


    Dim LngRetunCod         As Long     'メソッド戻り値
    Dim LngZoom             As Integer  '拡大縮小率

On Error GoTo ErrorHandle

    Output_SettingErr = False

    '初期化処理
    LngRetunCod = iObjCrDraw.Initialize
    If LngRetunCod <> 0 Then
        Output_SettingErr = False
        oStrErrMsg = "プリンタ初期化エラー(" & sGetErrMessage(LngRetunCod) & ")"
        Exit Function
    End If

    '用紙向き設定
    Select Case mStrOutPutOrient
    Case "Tate"
        iObjCrDraw.Printer.Orientation = corOrientVertical
    Case Else
        iObjCrDraw.Printer.Orientation = corOrientHorizontal
    End Select

    '用紙サイズ設定
    Select Case mStrOutPutSize
    Case "A3": iObjCrDraw.Printer.PaperSize = corPaperA3
    Case "B4": iObjCrDraw.Printer.PaperSize = corPaperB4
    Case "A4": iObjCrDraw.Printer.PaperSize = corPaperA4
    Case "B5": iObjCrDraw.Printer.PaperSize = corPaperB5
    Case "A5": iObjCrDraw.Printer.PaperSize = corPaperA5
    Case Else: iObjCrDraw.Printer.PaperSize = corPaperA4  '上記以外は A4 に設定
    End Select

    Select Case iStrPType
    Case "Print"
        'プリンタ出力開始処理
        If iObjCrDraw.CrPrinters.Count = 0 Then
            Output_SettingErr = False
            oStrErrMsg = "プリンタが未設定です"
            Exit Function
        End If

        '拡大縮小率を設定
        iObjCrDraw.Printer.PrintScale = mStrOutPutZoom

        'プリンタ出力を設定
        LngRetunCod = iObjCrDraw.StartPrinter(iStrPrinterNM, mStrCrfDocNameErr)

        '用紙向き設定
        Select Case mStrOutPutOrient
          Case "Tate"
            iObjCrDraw.Printer.Orientation = corOrientVertical
          Case Else
            iObjCrDraw.Printer.Orientation = corOrientHorizontal
         End Select

        '用紙サイズ設定
        Select Case mStrOutPutSize
         Case "A3": iObjCrDraw.Printer.PaperSize = corPaperA3
         Case "B4": iObjCrDraw.Printer.PaperSize = corPaperB4
         Case "A4": iObjCrDraw.Printer.PaperSize = corPaperA4
         Case "B5": iObjCrDraw.Printer.PaperSize = corPaperB5
         Case "A5": iObjCrDraw.Printer.PaperSize = corPaperA5
         Case Else: iObjCrDraw.Printer.PaperSize = corPaperA4  '上記以外は A4 に設定
        End Select

    End Select

    'エラー判定
    If LngRetunCod <> 0 Then
        Output_SettingErr = False
        oStrErrMsg = "プリンタ出力先設定エラー(" & sGetErrMessage(LngRetunCod) & ")"
        Exit Function
    End If

    '座標系単位指定
    iObjCrDraw.Unit = corUnitMm

    Output_SettingErr = True

    Exit Function

ErrorHandle:

    'ｴﾗｰ ﾒｯｾｰｼﾞ取得
    oStrErrMsg = ErrMsg_VB

End Function
'
'
'*-----------------------------------------------------------------------------*
'* 機能：ヘッダ項目出力処理
'* 概要：
'* 引数：ioObjCrForm       　　　　 Coフォームオブジェクト
'*
'* 戻値：（なし）
'*
'*-----------------------------------------------------------------------------*'
Private Sub Output_Head(ByRef ioObjCrForm As CrForm)

    With ioObjCrForm
        .CrObjects("発行日").text = Format(Now, "YYYY/MM/DD hh:nn:ss")
        .CrObjects("ページ").text = mLngPageCount
        '.CrObjects("Label9").Visible = False
    End With
    
End Sub

'*-----------------------------------------------------------------------------*
'* 機能：帳票出力終了設定
'* 概要：
'* 引数：iObjCrDraw       　　　　  Coreportsオブジェクト
'* 　　　iStrPType       　　　　   Print：印刷、View：プレビュ－表示
'* 　　　oStrErrMsg     　　　　    エラーメッセージ
'*
'* 戻値：Output_Setting         TRUE：正常、FALSE：エラー
'*
'*-----------------------------------------------------------------------------*'
Function Output_End(ByRef iObjCrDraw As CrDraw, _
                    ByVal iStrPType As String, _
                    ByRef oStrErrMsg As String) As Boolean


    Dim LngRetunCod            As Long     '戻り値
    Dim I               As Integer


On Error GoTo ErrorHandle
    
    Output_End = False

    Select Case iStrPType
    Case "Print"
        'プリンタ出力終了処理（印刷）
        LngRetunCod = iObjCrDraw.EndPrinter
        If LngRetunCod <> 0 Then
            Output_End = False
            oStrErrMsg = "プリンタ出力終了エラー(" & sGetErrMessage(LngRetunCod) & ")"
            Exit Function
        End If
        
    Case Else
        'プレビュー出力終了処理
        For I = 1 To mLngPageCount
            PreviewForm.Picture2(mLngPageCount).Refresh
        Next I
        PreviewForm.Caption = mStrCrfPvFrmName
    
    End Select

    Output_End = True
    
    Exit Function
    
    
ErrorHandle:
    
    'ｴﾗｰ ﾒｯｾｰｼﾞ取得
    oStrErrMsg = ErrMsg_VB
    
End Function

'--------------------------------------------------------------------------------
' 関数名     sGetErrMessage
' 概要
'           エラーメッセージ取得
' 引数
'           lRetval :メソッドの戻り値
' 戻り値
'           エラーメッセージ
' 機能説明
'           シーオーリポーツのエラーメッセージ取得関数
' 備考
'           なし  vbExclamation
'--------------------------------------------------------------------------------
Function sGetErrMessage(ByVal lRetval As Long) As String
    
    
    'メソッドの戻り値によってエラーメッセージをかえます。
    Select Case lRetval
        Case corNoError                  '  0
            sGetErrMessage = "正常終了"
        Case corErrNotOpenDocument       ' -1
            sGetErrMessage = "シーオーリポーツドキュメントファイル未オープン"
        Case corErrAlreadyOpenDocument   ' -2
            sGetErrMessage = "シーオーリポーツドキュメントファイルオープン済み"
        Case corErrOpenDocument          ' -3
            sGetErrMessage = "シーオーリポーツドキュメントファイルオープンエラー"
        Case corErrWriteDocument         ' -4
            sGetErrMessage = "シーオーリポーツドキュメントファイルライトエラー"
        Case corErrReadDocument          ' -5
            sGetErrMessage = "シーオーリポーツドキュメントファイルリードエラー"
        Case corErrNotOpenPrinter        ' -6
            sGetErrMessage = "プリンタ開始指定がない"
        Case corErrAlreadyOpenPrinter    ' -7
            sGetErrMessage = "プリンタ出力指定済み"
        Case corErrInvalidHdc            ' -8
            sGetErrMessage = "デバイスコンテキストが有効でない"
        Case corErrNoPrinter             ' -9
            sGetErrMessage = "プリンタが1つもインストールされていない"
        Case corErrPrinterNotExists      '-10
            sGetErrMessage = "指定されたプリンタがない"
        Case corErrCreateDc              '-11
            sGetErrMessage = "プリンタデバイスコンテキスト作成失敗"
        Case corErrStartDoc              '-12
            sGetErrMessage = "印刷JOB開始エラー"
        Case corErrStartPage             '-13
            sGetErrMessage = "ページ開始処理エラー"
        Case corErrEndPage               '-14
            sGetErrMessage = "ページ終了処理エラー"
        Case corErrEndDoc                '-15
            sGetErrMessage = "印刷JOB終了処理エラー"
        Case corErrCreateDocument        '-16
            sGetErrMessage = "シーオーリポーツドキュメントファイル作成エラー"
        Case corErrOleDraw               '-17
            sGetErrMessage = "OLE オブジェクトの描画に失敗"
        Case corErrPrinterControl        '-18
            sGetErrMessage = "プリンタ制御が正常に行われなかった"
        Case corErrOutputNone            '-19
            sGetErrMessage = "プリンタ制御エラー"
        Case corErrObjectDraw            '-20
            sGetErrMessage = "描画オブジェクトの描画に失敗"
        Case Else
            sGetErrMessage = Err.Description
    End Select


End Function

'==============================================================================*
' プロシージャ名　　：Set_Kaikou
' 機能　　　　　　　：改ページ処理
' 引数　　　　　　　：iObjCoForm        CRフォームオブジェクト
'                     iStrOutType       出力タイプ（Print：印刷、View：プレビュー）
'                     oStrErrMsg        エラーメッセージ
' 戻り値　　　　　　：Set_Kaikou
'                       TRUE :正常終了
'                       FALSE:エラー
' 作成日／作成者　　： 2003/ 6/ 20
' 備考　　　　　　　：
'------------------------------------------------------------------------------*
Function Set_Kaikou(ByRef iObjCoForm As CrForm, _
                    ByVal iStrOutType As String, _
                    ByVal Row As Long, _
                    ByRef oStrErrMsg As String) As Boolean

    Dim LngRetunCod     As Long         'リターンコード

On Error GoTo ERR_PROC

    Set_Kaikou = False
    
    'ページ数インクリメント
    mLngPageCount = mLngPageCount + 1
    
    If Row <> 0 Then
        'フォームオブジェクト出力処理
        LngRetunCod = iObjCoForm.PrintOut
    End If
    
    'エラー判定
    If LngRetunCod <> 0 Then
        Set_Kaikou = False
        oStrErrMsg = "データ出力エラー(" & sGetErrMessage(LngRetunCod) & ")"
        Exit Function
    End If
    
    'プレビューの時改ページ処理
    If iStrOutType = "View" Then
'        If Not Output_NewPage(oStrErrMsg) Then
'            Set_Kaikou = False
'            oStrErrMsg = oStrErrMsg & "[Output_NewPage]"
'            Exit Function
'        End If
    End If
        
    Set_Kaikou = True
        
    Exit Function

ERR_PROC:
    
    oStrErrMsg = ErrMsg_VB
    
End Function

'==============================================================================*
' プロシージャ名　　：Set_Kaikou
' 機能　　　　　　　：品番が存在しないエラー出力用の改ページ処理
' 引数　　　　　　　：iObjCoForm        CRフォームオブジェクト
'                     iStrOutType       出力タイプ（Print：印刷、View：プレビュー）
'                     oStrErrMsg        エラーメッセージ
' 戻り値　　　　　　：Set_Kaikou
'                       TRUE :正常終了
'                       FALSE:エラー
' 作成日／作成者　　： 2003/ 6/ 20
' 備考　　　　　　　：
'------------------------------------------------------------------------------*
Function Set_KaikouErr(ByRef iObjCoForm As CrForm, _
                    ByVal iStrOutType As String, _
                    ByVal Row As Long, _
                    ByRef oStrErrMsg As String) As Boolean

    Dim LngRetunCod     As Long         'リターンコード

On Error GoTo ERR_PROC

    Set_KaikouErr = False
    
    'ページ数インクリメント
    mLngPageCountErr = mLngPageCountErr + 1
    
    If Row <> 0 Then
        'フォームオブジェクト出力処理
        LngRetunCod = iObjCoForm.PrintOut
    End If
    
    'エラー判定
    If LngRetunCod <> 0 Then
        Set_KaikouErr = False
        oStrErrMsg = "データ出力エラー(" & sGetErrMessage(LngRetunCod) & ")"
        Exit Function
    End If
    
    'プレビューの時改ページ処理
    If iStrOutType = "View" Then
'        If Not Output_NewPage(oStrErrMsg) Then
'            Set_Kaikou = False
'            oStrErrMsg = oStrErrMsg & "[Output_NewPage]"
'            Exit Function
'        End If
    End If
        
    Set_KaikouErr = True
        
    Exit Function

ERR_PROC:
    
    oStrErrMsg = ErrMsg_VB
    
End Function
'*-----------------------------------------------------------------------------*
'* 機能：プリンタダイアログ表示
'*
'* 概要：           ：端末のプリンタ名称一覧の取得を行い、コンボボックスに設置する。
'*
'* 引数：           ：iObjCrDraw        CoReportsオブジェクト
'*                  ：oObjCmbPrinter    名称設定先コンボボックス
'*-----------------------------------------------------------------------------*
Public Sub Show_PrinterDialog(ByRef iObjCrDraw As CrDraw, _
                              ByRef oObjCmbPrinter As ComboBox)


    Dim ObjCrPrinter        As CrPrinter    'プリンタオブジェクト
    Dim LngDefPrint         As Long         '標準プリンタのインデックス
    Dim I                   As Long         'カウンタ
        
    oObjCmbPrinter.Clear

    With iObjCrDraw
        For Each ObjCrPrinter In .CrPrinters
            oObjCmbPrinter.AddItem ObjCrPrinter.PrinterName, I
            If ObjCrPrinter.PrinterName = .Printer.PrinterName Then
                LngDefPrint = I
            End If
            I = I + 1
        Next
        
        '標準プリンタを表示します。
        oObjCmbPrinter.ListIndex = LngDefPrint
    
    End With

End Sub

'==============================================================================*
' プロシージャ名　　：OutPut_Main
' 機能　　　　　　　：Ｃｒフォームオブジェクトへのデータセット処理
' 引数　　　　　　　：ioObjCrForm       Ｃｒオブジェクト
'                     iStrShoriYM       出力対象年月((YYYYMM)'
'                     iStrOutType       出力タイプ（Print：印刷、View：プレビュー）
'                     oStrErrMsg        エラーメッセージ
' 戻り値　　　　　　：OutPut_Main
'                       TRUE :正常終了
'                       FALSE:エラー
' 作成日／作成者　　：
' 備考　　　　　　　：
'
'------------------------------------------------------------------------------*
Private Function Output_Main(ByRef ioObjCrForm As CrForm, _
                             ByVal iStrOutType As String, _
                             ByVal LngOutRow As Long, _
                             ByRef oStrErrMsg As String) As Boolean

    Dim LngOutReturn        As Long                 '出力処理戻り値
    
'    Dim LngOutRow           As Long                 '出力行カウンタ
   
    Dim I                   As Long                 'カウンタ
   
    Dim intSTARTFLG As Integer
    
   
    
On Error GoTo ERR_PROC:

    Output_Main = False
     
    With ioObjCrForm
              'ページ印字
            If LngOutRow = 0 Then
                'ヘッダ部出力
                Call Output_Head(ioObjCrForm)
            End If
        'データの配置
                
                With JNW
                    '品番
                    If .品番 <> "" Then
                       ioObjCrForm.CrObjects("品番").ListText(0, LngOutRow) = .品番
                    End If
                    '背番号
                    If .背番号 <> "" Then
                       ioObjCrForm.CrObjects("背番号").ListText(0, LngOutRow) = .背番号
                    End If
                    '受入
                    If .受入 <> "" Then
                       ioObjCrForm.CrObjects("受入").ListText(0, LngOutRow) = .受入
                    End If
                    '納入番号
                    If .納入番号 <> "" Then
                       ioObjCrForm.CrObjects("納入番号").ListText(0, LngOutRow) = .納入番号
                    End If
                    '納入書行
                    If CStr(.納品書行) <> "" Then
                       ioObjCrForm.CrObjects("納品書行").ListText(0, LngOutRow) = CStr(.納品書行)
                    End If
                    '枚数
                    If CStr(.枚数) <> "" Then
                       ioObjCrForm.CrObjects("枚数").ListText(0, LngOutRow) = CStr(.枚数)
                    End If
                    '収容数
                    If CStr(.収容数) <> "" Then
                       ioObjCrForm.CrObjects("収容数").ListText(0, LngOutRow) = CStr(.収容数)
                    End If
                    '納入数
                    If CStr(.納入数) <> "" Then
                       ioObjCrForm.CrObjects("納入数").ListText(0, LngOutRow) = CStr(.納入数)
                    End If
                    '納入日
                    If .納入日 <> "" Then
                       ioObjCrForm.CrObjects("納入日").ListText(0, LngOutRow) = Format$(.納入日, "&&&&/&&/&&")
                    End If
                    '納入便
                    If .納入便 <> "" Then
                       ioObjCrForm.CrObjects("納入便").ListText(0, LngOutRow) = .納入便
                    End If
                    '納入先
                    If .納入先 <> "" Then
                       ioObjCrForm.CrObjects("納入先").ListText(0, LngOutRow) = .納入先
                    End If
                    '受注日
                    If .受注日 <> "" Then
                       ioObjCrForm.CrObjects("受注日").ListText(0, LngOutRow) = Format$(.受注日, "&&&&/&&/&&")
                    End If
                End With
                
                '改ページ処理
                If LngOutRow >= mLngPageMax - 1 Then
                   If Set_Kaikou(ioObjCrForm, iStrOutType, LngOutRow, oStrErrMsg) = False Then
                      Exit Function
                   Else
                      LngOutRow = 0
                       'データのクリア
                      ioObjCrForm.ClearAllFields
                       'ヘッダ部出力
                       Call Output_Head(ioObjCrForm)
                  End If
                End If
                
                 
'            ObjADORset.MoveNext
'        Loop
'
                 
    End With
    
         
    'フォームオブジェクト出力処理
'    LngOutReturn = ioObjCrForm.PrintOut
'    If LngOutReturn <> 0 Then
'        Output_Main = False
'        oStrErrMsg = "データ出力エラー(" & sGetErrMessage(LngOutReturn) & ")"
'        Exit Function
'    End If
    
    Output_Main = True
    
    Exit Function

ERR_PROC:
    
    oStrErrMsg = ErrMsg_VB

End Function

'==============================================================================*
' プロシージャ名　　：OutPut_MainErr
' 機能　　　　　　　：品番が存在しないエラー出力用
' 引数　　　　　　　：ioObjCrForm       Ｃｒオブジェクト
'                     iStrShoriYM       出力対象年月((YYYYMM)'
'                     iStrOutType       出力タイプ（Print：印刷、View：プレビュー）
'                     oStrErrMsg        エラーメッセージ
' 戻り値　　　　　　：OutPut_MainErr
'                       TRUE :正常終了
'                       FALSE:エラー
' 作成日／作成者　　：2004/07/26
' 備考　　　　　　　：
'
'------------------------------------------------------------------------------*
Private Function Output_MainErr(ByRef ioObjCrForm As CrForm, _
                             ByVal iStrOutType As String, _
                             ByVal LngOutRowErr As Long, _
                             ByRef oStrErrMsg As String) As Boolean

    Dim LngOutReturn        As Long                 '出力処理戻り値
    
'    Dim LngOutRow           As Long                 '出力行カウンタ
   
    Dim I                   As Long                 'カウンタ
   
    Dim intSTARTFLG As Integer
    
   
    
On Error GoTo ERR_PROC:

    Output_MainErr = False
     
    With ioObjCrForm
              'ページ印字
            If LngOutRowErr = 0 Then
                'ヘッダ部出力
'                Call Output_Head(ioObjCrForm)
                With ioObjCrForm
                    .CrObjects("発行日").text = Format(Now, "YYYY/MM/DD hh:nn:ss")
                    .CrObjects("ページ").text = mLngPageCountErr
                    '.CrObjects("Label9").Visible = False
                End With
            End If
        'データの配置
                
                With JNW
                    '品番
                    If .品番 <> "" Then
                       ioObjCrForm.CrObjects("品番").ListText(0, LngOutRowErr) = .品番
                    End If
                    '背番号
                    If .背番号 <> "" Then
                       ioObjCrForm.CrObjects("背番号").ListText(0, LngOutRowErr) = .背番号
                    End If
                    '受入
                    If .受入 <> "" Then
                       ioObjCrForm.CrObjects("受入").ListText(0, LngOutRowErr) = .受入
                    End If
                    '納入番号
                    If .納入番号 <> "" Then
                       ioObjCrForm.CrObjects("納入番号").ListText(0, LngOutRowErr) = .納入番号
                    End If
                    '納入書行
                    If CStr(.納品書行) <> "" Then
                       ioObjCrForm.CrObjects("納品書行").ListText(0, LngOutRowErr) = CStr(.納品書行)
                    End If
                    '枚数
                    If CStr(.枚数) <> "" Then
                       ioObjCrForm.CrObjects("枚数").ListText(0, LngOutRowErr) = CStr(.枚数)
                    End If
                    '収容数
                    If CStr(.収容数) <> "" Then
                       ioObjCrForm.CrObjects("収容数").ListText(0, LngOutRowErr) = CStr(.収容数)
                    End If
                    '納入数
                    If CStr(.納入数) <> "" Then
                       ioObjCrForm.CrObjects("納入数").ListText(0, LngOutRowErr) = CStr(.納入数)
                    End If
                    '納入日
                    If .納入日 <> "" Then
                       ioObjCrForm.CrObjects("納入日").ListText(0, LngOutRowErr) = Format$(.納入日, "&&&&/&&/&&")
                    End If
                    '納入便
                    If .納入便 <> "" Then
                       ioObjCrForm.CrObjects("納入便").ListText(0, LngOutRowErr) = .納入便
                    End If
                    '納入先
                    If .納入先 <> "" Then
                       ioObjCrForm.CrObjects("納入先").ListText(0, LngOutRowErr) = .納入先
                    End If
                    '受注日
                    If .受注日 <> "" Then
                       ioObjCrForm.CrObjects("受注日").ListText(0, LngOutRowErr) = Format$(.受注日, "&&&&/&&/&&")
                    End If
'*---------------------------------------------------<< 2004/08/31 >> Insert Start By CIS
                    '発注区分
                    If CisFun.Mid2(sDataOfLine, 3, 1) <> "" Then
                       ioObjCrForm.CrObjects("発注区分").ListText(0, LngOutRowErr) = CisFun.Mid2(sDataOfLine, 3, 1)
                    End If
'*---------------------------------------------------<< 2004/08/31 >> Insert End   By CIS
                End With
                
                '改ページ処理
                If LngOutRowErr >= mLngPageMax - 1 Then
                   If Set_KaikouErr(ioObjCrForm, iStrOutType, LngOutRowErr, oStrErrMsg) = False Then
                      Exit Function
                   Else
                      LngOutRowErr = 0
                       'データのクリア
                      ioObjCrForm.ClearAllFields
                       'ヘッダ部出力
                       With ioObjCrForm
                            .CrObjects("発行日").text = Format(Now, "YYYY/MM/DD hh:nn:ss")
                            .CrObjects("ページ").text = mLngPageCountErr
                        End With
                  End If
                End If
                 
    End With
    
    Output_MainErr = True
    
    Exit Function

ERR_PROC:
    
    oStrErrMsg = ErrMsg_VB

End Function

'+----------------------+
'+    DB内容ｾｯﾄ 処理     +
'+----------------------+
Private Sub DBSet_Sy()
'   <><><> 名称マスタ参照 <><><>
    Call ItemsClearSYM
    With SYM
        .区分名称 = "フタバ納入先管理"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "納入先"
        .英数字1 = "100"
    End With
    Call SYMReadChk
    If RTrim(SYM.英数字1) = "" Then SYM.英数字1 = "100"
    
    Call ItemsClearJNW
    
    With CisDB
    '   <><><> 受注管理NOの獲得 <><><>
        .SQL = "番号獲得"
        .StoadoCount = 3
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = OutPut用
        .ParaNo = 3: .ParaIO = Return用
    '
        .ParaNo = 1: .ParaValue = 11
    '
        If Not .DBStored Then
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(2) = "    受注管理NOの獲得に失敗しました            "
            CisFun.MB_Title = "採番処理"
            CisFun.MB_Button = Error
            CisFun.MBOX
            Exit Sub
        Else
            .ParaNo = 2: JNW.受注管理NO = .ParaValue
        End If
    End With
    
    With JNW
        .端末番号 = "' + substring(host_name(),1,20) + '"
        .受注区分 = 1
    '--↓受注日を受信データ上の発注日とする 2004/07/09 by TOTEC------------------------
        '''.受注日 = Format(Now, "YYYYMMDD")
        .受注日 = "20" & RTrim(CisFun.Mid2(sDataOfLine, 114, 6))
    '--↑受注日を受信データ上の発注日とする 2004/07/09 by TOTEC------------------------
        .納入日 = "20" & RTrim(CisFun.Mid2(sDataOfLine, 66, 6))
        .納入便 = RTrim(CisFun.Mid2(sDataOfLine, 72, 2))
        .納入先 = RTrim(SYM.英数字1) & (CisFun.Mid2(sDataOfLine, 100, 1))
        .出荷予定日 = .納入日
        .出荷予定便 = .納入便
'*---------------------------------------------------<< 2004/08/31 >> Update Start By CIS
'*      品番取引先マスタの受入をセット
'       .受入 = RTrim(CisFun.Mid2(sDataOfLine, 28, 2))
        .受入 = HTM.受入
'*---------------------------------------------------<< 2004/08/31 >> Update End   By CIS
        .納入番号 = RTrim(CisFun.Mid2(sDataOfLine, 60, 5))
        .納品書行 = CisFun.Val2(RTrim(CisFun.Mid2(sDataOfLine, 65, 1)))
        .品番 = RTrim(CisFun.Mid2(sDataOfLine, 35, 20))
        .背番号 = RTrim(CisFun.Mid2(sDataOfLine, 31, 4))
        .枚数 = CisFun.Val2(RTrim(CisFun.Mid2(sDataOfLine, 87, 3)))
        .収容数 = CisFun.Val2(RTrim(CisFun.Mid2(sDataOfLine, 55, 5)))
        .納入数 = CisFun.Val2(RTrim(CisFun.Mid2(sDataOfLine, 90, 6)))
        .作成種別 = "F"
    
'*---------------------------------------------------<< 2004/09/09 >> Insert Start By CIS
        If HTM.伝票種類 = "" Then .納品書発行区分 = "9"
'*---------------------------------------------------<< 2004/09/09 >> Insert End   By CIS
    
    End With

    Call JNWInsert
    
End Sub
'<><><> Body部 <><><>
Private Function DBSet_Chk() As Boolean
    DBSet_Chk = False
'   <><><> 名称マスタ参照 <><><>
    Call ItemsClearSYM
    With SYM
        .区分名称 = "フタバ納入先管理"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "納入先"
        .英数字1 = "100"
    End With
    Call SYMReadChk
    If RTrim(SYM.英数字1) = "" Then SYM.英数字1 = "100"
    
    gSL_Select = "SELECT * FROM 受注入力ワーク "
    gSL_Select = gSL_Select & " WHERE 納入日 = '" & "20" & RTrim(CisFun.Mid2(sDataOfLine, 66, 6)) & "'"
    gSL_Select = gSL_Select & " AND 納入便 = '" & RTrim(CisFun.Mid2(sDataOfLine, 72, 2)) & "'"
    gSL_Select = gSL_Select & " AND 納入番号 = '" & RTrim(CisFun.Mid2(sDataOfLine, 60, 5)) & "'"
    gSL_Select = gSL_Select & " AND 品番 = '" & RTrim(CisFun.Mid2(sDataOfLine, 35, 20)) & "'"
    gSL_Select = gSL_Select & " AND 納入先 = '" & RTrim(SYM.英数字1) & (CisFun.Mid2(sDataOfLine, 100, 1)) & "'"
    If JNWRead(gSL_Select, 1) Then
       nCountErr = nCountErr + 1
       Exit Function
    End If

DBSet_Chk_Err:
    DBSet_Chk = True
End Function

'*****************************************************************
'2004/07/26
'品番が存在するかチェックする
'*****************************************************************
Private Function Chk_Hin() As Boolean
    Chk_Hin = False
    
    gSL_Select = ""
    gSL_Select = "SELECT * FROM 品番マスタ "
    gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(CisFun.Mid2(sDataOfLine, 35, 20)) & "'"
    
    CisDB.SQL = gSL_Select
    If Not CisDB.DBRead(1, 0) Then
'        MsgBox "aaa"
        Exit Function
    End If
    Call JNWClose
    
''''''    If Not JNWRead(gSL_Select, 1) Then
'''''''       nCountErr = nCountErr + 1
''''''        MsgBox "AAAAAAAAAa"
''''''       Exit Function
''''''    End If
    '構造体に値をセット 2004/07/30
    With SYM
        .区分名称 = "フタバ納入先管理"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "納入先"
        .英数字1 = "100"
    End With
'*---------------------------------------------------<< 2004/08/31 >> Insert Start By CIS
    Call SYMReadChk
    If RTrim(SYM.英数字1) = "" Then SYM.英数字1 = "100"
'*---------------------------------------------------<< 2004/08/31 >> Insert End   By CIS
    
'*---------------------------------------------------<< 2004/08/31 >> Update Start By CIS
'   gSL_Select = ""
'   gSL_Select = "SELECT * FROM 品番取引先マスタ "
'   gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(CisFun.Mid2(sDataOfLine, 35, 20)) & "'"
'   gSL_Select = gSL_Select & " AND 取引先 = '" & RTrim(SYM.英数字1) & (CisFun.Mid2(sDataOfLine, 100, 1)) & "'"
'
'   CisDB.SQL = gSL_Select
'   If Not CisDB.DBRead(1, 0) Then
''        MsgBox "aaa"
'       Exit Function
'   End If
'   Call JNWClose
'
'*  SK区分の追加：発注区分（３桁目）= 1 Then かんばん Else 指示
    gSL_Select = ""
    
    
'*---------------------------------------------------<< 2004/09/09 >> Update Start By CIS
'    gSL_Select = "SELECT 受入 FROM 品番取引先マスタ "
    gSL_Select = "SELECT 受入, 伝票種類 FROM 品番取引先マスタ "
'*---------------------------------------------------<< 2004/09/09 >> Update End   By CIS
    
    gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(CisFun.Mid2(sDataOfLine, 35, 20)) & "'"
    gSL_Select = gSL_Select & " AND 取引先 = '" & RTrim(SYM.英数字1) & (CisFun.Mid2(sDataOfLine, 100, 1)) & "'"
    Select Case CisFun.Mid2(sDataOfLine, 3, 1)
        Case "1":   gSL_Select = gSL_Select & " AND SK区分 = 0"
        Case Else:  gSL_Select = gSL_Select & " AND SK区分 = 1"
    End Select
    If Not HTMRead(gSL_Select, 1, 0) Then Exit Function
'*---------------------------------------------------<< 2004/08/31 >> Update End   By CIS


Chk_Hin_Err:
    Chk_Hin = True
End Function
'****************************************************************
'*　受注入力ワーク(未累積データ)に登録済みで、データ区分が"C"の場合は、
'*　受注入力ワークの該当データを削除する
'*
'*　2004/07/08      追加　by TOTEC
'****************************************************************
Private Function DBSet_Del(Optional DBNo As Integer) As Boolean
    DBSet_Del = False
'   <><><> 名称マスタ参照 <><><>
    Call ItemsClearSYM
    With SYM
        .区分名称 = "フタバ納入先管理"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "納入先"
        .英数字1 = "100"
    End With
    Call SYMReadChk
    If RTrim(SYM.英数字1) = "" Then SYM.英数字1 = "100"
    
    gSL_Select = "DELETE FROM 受注入力ワーク "
    gSL_Select = gSL_Select & " WHERE 納入日 = '" & "20" & RTrim(CisFun.Mid2(sDataOfLine, 66, 6)) & "'"
    gSL_Select = gSL_Select & " AND 納入便 = '" & RTrim(CisFun.Mid2(sDataOfLine, 72, 2)) & "'"
    gSL_Select = gSL_Select & " AND 納入番号 = '" & RTrim(CisFun.Mid2(sDataOfLine, 60, 5)) & "'"
    gSL_Select = gSL_Select & " AND 品番 = '" & RTrim(CisFun.Mid2(sDataOfLine, 35, 20)) & "'"
    gSL_Select = gSL_Select & " AND 納入先 = '" & RTrim(SYM.英数字1) & (CisFun.Mid2(sDataOfLine, 100, 1)) & "'"
    gSL_Select = gSL_Select & " AND 累積区分 = 0"
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With

DBSet_Chk_Err:
    DBSet_Del = True
End Function
'****************************************************************
'*　受注入力ワーク(未累積データ)に登録済みで、データ区分が"A"の場合は、
'*　受注入力ワークの該当データを上書きする
'*
'*　2004/07/08      追加　by TOTEC
'****************************************************************
Private Function DBSet_Upd(Optional DBNo As Integer) As Boolean
    DBSet_Upd = False
'   <><><> 名称マスタ参照 <><><>
    Call ItemsClearSYM
    With SYM
        .区分名称 = "フタバ納入先管理"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "納入先"
        .英数字1 = "100"
    End With
    Call SYMReadChk
    If RTrim(SYM.英数字1) = "" Then SYM.英数字1 = "100"
    
    With JNW
        gSL_Select = "Update 受注入力ワーク Set"
        gSL_Select = gSL_Select & " 端末番号='" & RTrim(.端末番号) & "',"
        gSL_Select = gSL_Select & " 受注区分=" & .受注区分 & ","
        gSL_Select = gSL_Select & " 受注日='" & RTrim(.受注日) & "',"
        gSL_Select = gSL_Select & " 納入日='" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & " 納入便='" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & " 納入先='" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & " 出荷予定日='" & RTrim(.出荷予定日) & "',"
        gSL_Select = gSL_Select & " 出荷予定便='" & RTrim(.出荷予定便) & "',"
'*---------------------------------------------------<< 2004/08/31 >> Update Start By CIS
'*      品番取引先マスタの受入をセット
'       gSL_Select = gSL_Select & " 受入='" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & " 受入='" & RTrim(HTM.受入) & "',"
'*---------------------------------------------------<< 2004/08/31 >> Update End   By CIS
        gSL_Select = gSL_Select & " 納入番号='" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & " 納品書行=" & .納品書行 & ","
        gSL_Select = gSL_Select & " 品番='" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & " 背番号='" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & " 枚数=" & .枚数 & ","
        gSL_Select = gSL_Select & " 収容数=" & .収容数 & ","
        gSL_Select = gSL_Select & " 納入数=" & .納入数
        
        gSL_Select = gSL_Select & " WHERE 納入日 = '" & "20" & RTrim(CisFun.Mid2(sDataOfLine, 66, 6)) & "'"
        gSL_Select = gSL_Select & " AND 納入便 = '" & RTrim(CisFun.Mid2(sDataOfLine, 72, 2)) & "'"
        gSL_Select = gSL_Select & " AND 納入番号 = '" & RTrim(CisFun.Mid2(sDataOfLine, 60, 5)) & "'"
        gSL_Select = gSL_Select & " AND 品番 = '" & RTrim(CisFun.Mid2(sDataOfLine, 35, 20)) & "'"
        gSL_Select = gSL_Select & " AND 納入先 = '" & RTrim(SYM.英数字1) & (CisFun.Mid2(sDataOfLine, 100, 1)) & "'"
        gSL_Select = gSL_Select & " AND 累積区分 = 0"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With

DBSet_Chk_Err:
    DBSet_Upd = True
End Function
'□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□
'□
'□     初期設定の読込み＆書込み
'□
'□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□
'+---------------------------+
'+      初期設定の読込み
'+---------------------------+
Private Function Text_IniGet(FileName As String) As Boolean
    Text_IniGet = False
    With CisFun
        .INI_Section = "フタバ産業"
        .INI_Key = "受注テキスト"
        .INI_Default = ""
        .INI_FileID = gIniExe & gIniName
        .GetIni
        FileName = RTrim(.INI_String)
    End With
    Text_IniGet = True
End Function
'+---------------------------+
'+      初期設定の書込み
'+---------------------------+
Private Function Text_IniPut(FileName As String) As Boolean
    Text_IniPut = False
    With CisFun
        .INI_Section = "フタバ産業"
        .INI_Key = "受注テキスト"
        .INI_String = FileName
        .INI_FileID = gIniExe & gIniName
        .PutIni
    End With
    Text_IniPut = True
End Function
'******************************************************************
'*      初期処理
'******************************************************************
Private Function InitProc() As Boolean
    InitProc = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select *"
    gSL_Select = gSL_Select & "  from 受注入力ワーク"
    gSL_Select = gSL_Select & " where 端末番号 = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and 累積区分 = 0"
    If JNWRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 6
            .MB_MSG(1) = "    当端末処理分の受注未累積データが存在します。            "
            .MB_MSG(4) = "      『受注データ取込』処理を終了しますか？"
            .MB_MSG(5) = "      いいえの場合は、未累積データを削除し続行します"
            .MB_Title = "未累積データの確認"
            .MB_Button = Yes_No
            If .MBOX Then Exit Function
        End With
    End If

    With CisDB
        On Error GoTo InitProc_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
'       << 受注入力テーブル >>
        .SQL = ""
        .SQL = .SQL & "delete"
        .SQL = .SQL & "  from 受注入力ワーク"
        .SQL = .SQL & " where 端末番号 = substring(host_name(),1,20)"
        .SQL = .SQL & "   and 累積区分 = 0"
        .DBExec
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        On Error GoTo 0
    End With

    InitProc = True
    Exit Function
InitProc_Err:
    CisFun.ErrorBox
    End
End Function
'*---------------------------------------*'
'*      TEXT変換履歴テーブル 更新処理       *
'*---------------------------------------*'
Private Function TextRirekiPutX(THKNo As Long, PutKB As String, ErrSet As String, Optional DBNo As Integer) As Boolean

    TextRirekiPutX = False
    On Error GoTo TextRirekiPut_ER:
    
    If THKNo = 0 Then
        gSL_Select = "SELECT MAX(履歴管理番号) + 1 AS 履歴管理番号 FROM TEXT変換履歴テーブル "
        If THRRead(gSL_Select, 1, DBNo) Then
            THKNo = THR.履歴管理番号
        End If
    End If
    If THKNo = 0 Then
       THKNo = 1
    End If

    gSL_Select = "Select * From TEXT変換履歴テーブル"
    gSL_Select = gSL_Select & " Where 履歴管理番号 = " & THKNo
    If Not THRRead(gSL_Select, 1, DBNo) Then
        Call ItemsClearTHR
        With THR
            .履歴管理番号 = THKNo
            .開始日時 = Now()
            .FILECD = mFileCode
            .終了日時 = Now()
            .作成者 = gTanto
            .処理コード = 1
        End With
        Call THRInsert(1)
        GoTo TextRirekiPut_ED
    End If
    
    With THR
        .HEAD件数 = 0
        .TAIL件数 = 0
        If PutKB = "T" Then
           .件数 = CisFun.Val2(B1lb_CntGet.Caption)
        Else
           .件数 = nCountGet
        End If
        .変換件数 = .件数
        .文言件数 = 0
        .出力件数 = CisFun.Val2(B1lb_CntPut.Caption)
        If Trim(ErrSet) = "" Then
          .処理コード = 0
          .受信エラー = ""
        Else
          .処理コード = 1
          .受信エラー = ErrSet
        End If
        .終了日時 = Now()
    End With
    Call THRUpdate(1)
    
    If PutKB = "E" Then
       THKNo = 0
    End If
TextRirekiPut_ED:
    On Error GoTo 0
    TextRirekiPutX = True
    Exit Function
TextRirekiPut_ER:
    
End Function




