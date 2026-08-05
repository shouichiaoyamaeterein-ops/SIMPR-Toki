VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CXM0000 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "SIMPR-A 環境設定"
   ClientHeight    =   4320
   ClientLeft      =   3765
   ClientTop       =   2580
   ClientWidth     =   8550
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   4320
   ScaleWidth      =   8550
   Begin Cis3D_v60.CIS3D PB_Option 
      Height          =   315
      Left            =   6840
      Top             =   3945
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   556
      Caption         =   "オプション ( F1 )"
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
      cBoderWidth     =   3
      cAlingnment     =   7
      cButton         =   -1  'True
   End
   Begin Cis3D_v60.CIS3D B_UpDate 
      Height          =   315
      Left            =   1455
      Top             =   3945
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   556
      BackColor       =   8421504
      ForeColor       =   16777215
      Caption         =   "Update 9999/99/99"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   8421504
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cFont3D         =   2
      cAlingnment     =   1
      cBoderStyle     =   2
   End
   Begin Cis3D_v60.CIS3D B_Ver 
      Height          =   315
      Left            =   45
      Top             =   3945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BackColor       =   8421504
      ForeColor       =   16777215
      Caption         =   "Ver 1.00"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   8421504
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cFont3D         =   2
      cAlingnment     =   7
      cBoderStyle     =   2
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   4395
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   6135
      ScaleHeight     =   555
      ScaleWidth      =   2325
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   30
      Width           =   2385
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   555
         Left            =   1530
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   555
         Left            =   780
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnCaption      =   "更新"
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
      Begin CisBtn_60.CisBtn PB_Prt 
         Height          =   555
         Left            =   30
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
   Begin Cis3D_v60.CIS3D CIS3D2 
      Height          =   675
      Left            =   -30
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   1191
      ForeColor       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HGSｺﾞｼｯｸE"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   3
      cAlingnment     =   1
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   495
         Left            =   1485
         Top             =   90
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
         ForeColor       =   16711680
         Caption         =   "環   境   設   定"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "HGSｺﾞｼｯｸE"
            Size            =   21.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
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
   Begin VB.Timer Timer1 
      Left            =   6150
      Top             =   270
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   3165
      Left            =   45
      ScaleHeight     =   3105
      ScaleWidth      =   8400
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   735
      Width           =   8460
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   1500
         Left            =   60
         Top             =   45
         Width           =   8280
         _ExtentX        =   14605
         _ExtentY        =   2646
         BackColor       =   14737632
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin Cis3D_v60.CIS3D UC_3D3 
            Height          =   495
            Left            =   60
            Top             =   420
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   873
            BackColor       =   14737632
            Caption         =   "システム名"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cPositionX      =   50
            Begin CisText_V60.CisText iB1_SysName 
               Height          =   345
               Left            =   1200
               TabIndex        =   0
               Top             =   90
               Width           =   3405
               _ExtentX        =   6006
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
               MaxLength       =   30
               IMEMode         =   4
               cCaps           =   2
            End
         End
         Begin Cis3D_v60.CIS3D PFXLabel1 
            Height          =   405
            Index           =   0
            Left            =   15
            Top             =   0
            Width           =   8265
            _ExtentX        =   14579
            _ExtentY        =   714
            BackColor       =   4210752
            ForeColor       =   16777215
            Caption         =   "《 システム環境 》"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   14.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D UC_3D4 
            Height          =   495
            Left            =   60
            Top             =   930
            Width           =   8115
            _ExtentX        =   14314
            _ExtentY        =   873
            BackColor       =   14737632
            Caption         =   "ユーザｺｰﾄﾞ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
            Begin CisText_V60.CisText iB1_UserID 
               Height          =   345
               Left            =   1200
               TabIndex        =   3
               Top             =   90
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXX"
               MaxLength       =   6
               cCaps           =   2
            End
            Begin CisText_V60.CisText iB1_UserName 
               Height          =   345
               Left            =   4395
               TabIndex        =   4
               Top             =   90
               Width           =   3540
               _ExtentX        =   6244
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   30
               IMEMode         =   4
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D UC_3D5 
               Height          =   345
               Left            =   3420
               Top             =   90
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   609
               BackColor       =   14737632
               Caption         =   "ユーザ名"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   14737632
               cBoderColor2    =   14737632
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cFont3D         =   2
               cAlingnment     =   1
               cBoderStyle     =   2
               cPositionX      =   50
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   495
            Index           =   0
            Left            =   5070
            Top             =   420
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   873
            BackColor       =   14737632
            Caption         =   "Logo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
            Begin VB.CheckBox Ch_LogoDisp 
               BackColor       =   &H00E0E0E0&
               Caption         =   "非表示"
               Height          =   255
               Left            =   555
               TabIndex        =   1
               Top             =   135
               Width           =   855
            End
            Begin CisText_V60.CisText iB1_Logo 
               Height          =   345
               Left            =   1545
               TabIndex        =   2
               Top             =   75
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX1"
               MaxLength       =   10
               IMEMode         =   2
               cCaps           =   2
            End
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D6 
         Height          =   1515
         Left            =   60
         Top             =   1545
         Width           =   8280
         _ExtentX        =   14605
         _ExtentY        =   2672
         BackColor       =   14737632
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin Cis3D_v60.CIS3D UC_3D7 
            Height          =   495
            Left            =   60
            Top             =   450
            Width           =   5385
            _ExtentX        =   9499
            _ExtentY        =   873
            BackColor       =   14737632
            Caption         =   "実行ﾃﾞｨﾚｸﾄﾘ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cPositionX      =   50
            Begin CisText_V60.CisText iB1_ExeDir 
               Height          =   345
               Left            =   1200
               TabIndex        =   5
               Top             =   90
               Width           =   3795
               _ExtentX        =   6694
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   $"CXM0000.frx":0000
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_ExeDir 
               Height          =   315
               Left            =   4950
               Top             =   90
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               ForeColor       =   0
               Caption         =   "…"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cBoderWidth     =   3
               cAlingnment     =   8
            End
         End
         Begin Cis3D_v60.CIS3D PFXLabel1 
            Height          =   435
            Index           =   1
            Left            =   0
            Top             =   0
            Width           =   8265
            _ExtentX        =   14579
            _ExtentY        =   767
            BackColor       =   4210752
            ForeColor       =   16777215
            Caption         =   "《 ｸﾗｲｱﾝﾄ動作環境 》"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   14.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D UC_3D100 
            Height          =   495
            Index           =   0
            Left            =   60
            Top             =   960
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   873
            BackColor       =   14737632
            Caption         =   "サーバー名"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cBoderStyle     =   1
            cPositionX      =   50
            cPositionY      =   120
            Begin CisText_V60.CisText iB1_Server 
               Height          =   345
               Left            =   1200
               TabIndex        =   6
               Top             =   90
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   20
               IMEMode         =   2
               cCaps           =   2
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D100 
            Height          =   495
            Index           =   2
            Left            =   2805
            Top             =   960
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   873
            BackColor       =   14737632
            Caption         =   "ユーザーID"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cBoderStyle     =   1
            cPositionX      =   50
            cPositionY      =   120
            Begin CisText_V60.CisText iB1_UID 
               Height          =   345
               Left            =   1110
               TabIndex        =   7
               Top             =   90
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   20
               IMEMode         =   2
               cCaps           =   2
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D100 
            Height          =   495
            Index           =   3
            Left            =   5460
            Top             =   960
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   873
            BackColor       =   14737632
            Caption         =   "パスワード"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cBoderStyle     =   1
            cPositionX      =   50
            cPositionY      =   120
            Begin VB.TextBox iB1_Pass 
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               IMEMode         =   3  'ｵﾌ固定
               Left            =   1200
               MaxLength       =   15
               PasswordChar    =   "*"
               TabIndex        =   8
               Text            =   "*************"
               Top             =   90
               Width           =   1425
            End
         End
         Begin Cis3D_v60.CIS3D Back_DBName 
            Height          =   495
            Left            =   5460
            Top             =   450
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   873
            BackColor       =   14737632
            Caption         =   "ﾃﾞｰﾀﾍﾞｰｽ名"
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cFont3D         =   2
            cBoderStyle     =   1
            cPositionX      =   50
            cPositionY      =   120
            Begin CisText_V60.CisText iB1_DBName 
               Height          =   345
               Left            =   1200
               TabIndex        =   9
               Top             =   90
               Width           =   1450
               _ExtentX        =   2566
               _ExtentY        =   609
               cLostColor      =   16777152
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   20
               IMEMode         =   2
               cCaps           =   2
            End
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D10 
         Height          =   555
         Left            =   30
         Top             =   6045
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   979
         BackColor       =   14737632
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin Cis3D_v60.CIS3D UC_3D11 
            Height          =   435
            Left            =   2370
            Top             =   60
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   767
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
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
            cPositionX      =   50
            Begin VB.CommandButton cmdPass 
               Caption         =   "ﾊﾟｽﾜｰﾄﾞ変更(&P)"
               Height          =   375
               Left            =   30
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   30
               Width           =   1455
            End
         End
         Begin Cis3D_v60.CIS3D PFXLabel1 
            Height          =   555
            Index           =   2
            Left            =   0
            Top             =   0
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   979
            BackColor       =   4210752
            ForeColor       =   16777215
            Caption         =   "《 その他環境 》"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cFont3D         =   2
            cAlingnment     =   7
         End
      End
   End
End
Attribute VB_Name = "CXM0000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  環境設定
'**       フォームID    :  CXM0000
'**
'**       処理概要      :
'**
'**       作  成  日    :  2001/11/07  By CIS
'**       変  更  日    :  2004/09/24  By CIS - Logo情報は SIMPR-A.INI も更新
'**       変  更  日    :  2005/03/04  By CIS - 重複起動ﾁｪｯｸで起動済の時、『実行時エラー'91'』を表示する対応（Timerﾌﾟﾛﾊﾟﾃｨの Interval=0 に修正）
'**       変  更  日    :  2005/09/27  By CIS - ﾊﾟﾗﾒｰﾀｰ設定時、終了後ﾊﾟﾗﾒｰﾀｰの実行を行う。
'**                                             ﾊﾟｽﾜｰﾄﾞ設定時、更新ﾎﾞﾀﾝ押下ｴﾗｰとなる現象を修正
'**       変  更  日    :  2006/03/14  By CIS - ﾊﾟｽﾜｰﾄﾞ設定時、更新ﾎﾞﾀﾝ押下ｴﾗｰとなる現象を修正
'**       変  更  日    :  2006/04/10  By CIS - 印刷機能追加
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ

    Dim mKidoExe        As String
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   # 二 重 起 動 防 止 #
'   #------------------#
    If Not Dupli_Start Then End

    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me

    Timer1.Interval = 1000          ' 時刻表示間隔
    Timer1.Enabled = True           ' 時刻表示開始
    B1_Area1.BackColor = gPGotSel

'   #------------------#
'   # 初 期 内 容 取 得 #
'   #------------------#
    If Not IniGet Then End
'+-------------------------------------------- 2005/09/27 Insert
    If InStr(1, Command, "/") <> 0 Then
       mKidoExe = Trim(Command)
       mKidoExe = Replace(mKidoExe, "/", "")
    End If
'+-------------------------------------------- 2005/09/27 Insert End

'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Call CisFrm.InitFld

'   #------------------#
'   #   画面中央表示    #
'   #------------------#
    Call CisFrm.Frm_Center(False)
    
'   ( 内容表示 )
    Call BodyDisp

'   ( 組込起動 )
    If gEnvFlg = "" Then Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
'   ( タイトルバー用 )
    With CisFrm
        .TimeSet_TitleName = "SIMPR-A 環境設定"
        .TimeSet_TitleSpace = 40
    End With
    
    FormAct = True
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down") Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》       +
'+----------------------------------------------------------+
Function Key_Acc(Key_Code As Integer, IVENT As String)
    If Not FormAct Then Exit Function
    
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    
    Select Case Key_Code
        Case vbKeyF1:       Call PB_Option_Click                        ' オプション
        Case vbKeyF12:      If PB_ENT.Visible Then Call PB_END_Click    '【終了】
        Case vbKeyReturn:   If PB_ENT.Visible Then Call PB_ENT_Click    '【入力】
        Case Else:          Exit Function
    End Select
    
Key_Acc_ED:
    Key_Acc = True
End Function
'+------------------------+
'+   Ch_LogoDisp(ﾛｺﾞ表示)
'+------------------------+
Private Sub Ch_LogoDisp_Click()
    If Ch_LogoDisp.Value = 0 Then
        Ch_LogoDisp.Caption = "非表示"
        iB1_Logo.Enabled = False
    Else
        Ch_LogoDisp.Caption = "表示"
        iB1_Logo.Enabled = True
    End If
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
   Call ReturnPress
End Sub
'+---------------------+
'+    終了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    If gEnvFlg = "" Then
        Call CisDB.DBDISConnect
        Call CisFrm.UnLoadDisp
    '+---------------------------------------------- 2005/09/27 Insert
        On Error Resume Next
        If Trim(mKidoExe) <> "" Then
           Shell gIniExe & mKidoExe
        End If
    '+---------------------------------------------- 2005/09/27 Insert End
        End
    Else
        Me.Hide
    End If
End Sub
'---------------------------------------------- 2006/04/10 Insert
'+---------------------+
'+       印  刷        +
'+---------------------+
Private Sub PB_Prt_Click()
    
    CXM0005.Show vbModal
    Unload CXM0005
    Set CXM0005 = Nothing
    
    If CXM0000_PrtKB = 9 Then Exit Sub
        
    Call CXM0000LPrint
End Sub
Private Sub CIS3D3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 3 Then
       CXM0000_PrtKB = 3
       Call CXM0000LPrint
    End If
End Sub
'---------------------------------------------- 2006/04/10 Insert End
'+---------------------+
'+   オプションボタン   +
'+---------------------+
Private Sub PB_Option_Click()
    gRtn = True
    
    CSYSOPTION.Show vbModal
    Unload CSYSOPTION
    Set CSYSOPTION = Nothing
End Sub
'+---------------------+
'+   実行ﾃﾞｨﾚｸﾄﾘ　参照　+
'+---------------------+
Private Sub PB_ExeDir_Click()
    With CisFun
        .Folder_Title = "実行ディレクトリ参照"
        .Folder_DefDir = iB1_ExeDir
        .FolderSel
        
        If RTrim$(.Folder_Result) <> "" Then iB1_ExeDir = RTrim$(.Folder_Result)
    End With
    iB1_ExeDir.SetFocus
End Sub
'+---------------------+
'+   iH1_Pass
'+---------------------+
Private Sub iB1_Pass_GotFocus()
    iB1_Pass.BackColor = gIGotSel
    iB1_Pass.SelStart = 0
    iB1_Pass.IMEMode = iB1_Pass.IMEMode
    iB1_Pass.SelLength = iB1_Pass.MaxLength
End Sub
Private Sub iB1_Pass_LostFocus()
    iB1_Pass.BackColor = gILostSel
End Sub
Private Sub iB1_Pass_KeyPress(KeyAscii As Integer)
    iB1_Pass.SelLength = iB1_Pass.MaxLength
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If Not BodyChk Then GoTo ReturnPress_Ed
    If Not DBPut Then GoTo ReturnPress_Ed
    If gEnvFlg = "CALL" Then
       gEnvFlg = "OK"
    End If
ReturnPress_Ed:
    Call CisFrm.MousePT(1)
End Sub
'****************************
'*      BODYﾁｪｯｸ処理
'****************************
Private Function BodyChk() As Boolean
    Dim wSaveAcc        As String
    BodyChk = False
' ﾕｰｻﾞｰｺｰﾄﾞ入力ﾁｪｯｸ
    If Trim(iB1_UserID) = "" Then
       MsgBox "ユーザーコードを入力して下さい。", vbExclamation, "必須入力ｴﾗｰ"
       iB1_UserID.SetFocus
       Exit Function
    End If
    If Trim(iB1_UserName) = "" Then
       gLong = MsgBox("ユーザー名が入力されていません。       " & vbCrLf & vbCrLf & _
              "        よろしいですか？                      ", vbExclamation + vbYesNo, "入力確認")
       If gLong <> vbYes Then
          iB1_UserName.SetFocus
          Exit Function
       End If
    End If
    
    If Len(iB1_Pass) <> iB1_Pass.MaxLength Then
        iB1_Pass = RTrim$(iB1_Pass) & Space(iB1_Pass.MaxLength - Len(RTrim$(iB1_Pass)))
    End If
    
' データベース接続確認
    Call CisDB.DBDISConnect

    With CisDB
        .DBName = iB1_DBName                ' 接続データベース名
        .User = iB1_UID                     ' 接続ユーザー
    '+----------------------------------------------------------------- 2005/09/27 Delete
'        .PASS = RTrim(iB1_Pass)             ' 接続パスワード
    '+----------------------------------------------------------------- 2005/09/27 Delete End
    '+----------------------------------------------------------------- 2005/09/27 Insert
        .PASS = CisDB.PassWord(RTrim$(iB1_Pass))
    '+----------------------------------------------------------------- 2005/09/27 Insert End
        .Server = iB1_Server                ' 接続サーバー名
    End With

    '+----------------------------------------------------------------- 2006/03/14 Delete
    'If Not CisDB.DBConnect(False) Then
    '+----------------------------------------------------------------- 2006/03/14 Delete End
    '+----------------------------------------------------------------- 2006/03/14 Insert
    CisDB.PassLink
    If Not CisDB.DBConnect(App.Title) Then
    '+----------------------------------------------------------------- 2006/03/14 Insert End
        With CisFun
            .MB_Lines = 5
            .MB_MSG(1) = "   データベースに接続出来ません！！"
            .MB_MSG(3) = "   設定等を見直して再度実行して下さい　　"
            .MB_Title = "ﾃﾞｰﾀﾍﾞｰｽ接続"
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If

    BodyChk = True
End Function
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    Call CisFrm.TimeSet
End Sub
'+----------------------+
'+     内 容 表 示      +
'+----------------------+
Private Sub BodyDisp()
    Static wVersion As String
    Static wKeiyaku As String
    Static wObj     As CIS3D
'
    B_Ver = "Ver " & RTrim(gSysVer)
    B_UpDate = "UpDate " & RTrim(gSysUpDate)

'システム環境
    iB1_SysName = gSysName
    iB1_UserID = gCompany
    iB1_UserName = gCompName
    iB1_Logo = gLogo
    If gLogoDisp Then Ch_LogoDisp.Value = 1 Else Ch_LogoDisp.Value = 0
    Call Ch_LogoDisp_Click

'クライアント実行環境
    iB1_ExeDir = gIniExe
    iB1_DBName = CisDB.DBName
    iB1_Server = CisDB.Server
    iB1_UID = CisDB.User
    
    If gCipher = 1 Then gStr = RTrim$(CisDB.PASS)
    gStr = RTrim$(CisDB.PASS)
    gCipher = 0
    
    If gStr = "" Or gStr = "&&(%%$$(((" Then
       iB1_Pass = Space(iB1_Pass.MaxLength)
    Else
       iB1_Pass = gStr & Space(iB1_Pass.MaxLength - Len(gStr))
    End If
End Sub
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()

    DBPut = False
   Dim Atai As String
   Dim DIRECT As String
   Dim POS As Long
   Dim ALEN As Integer
   Dim APOS As Integer

'  ( 確認メッセージ )
    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = "     【 環境設定変更確認 】      "
        .MB_MSG(3) = "    環境の設定を変更しますか      "
        .MB_Title = "環境設定"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With

' 実行ﾃﾞｨﾚｸﾄﾘ
    SaveSetting appname:=gRegKey, section:="Environment", _
                        Key:="Exe", setting:=RTrim(iB1_ExeDir)
    
    If Right$(RTrim$(iB1_ExeDir), 1) <> "\" Then gStr = "\" Else gStr = ""
    gIniExe = RTrim$(iB1_ExeDir) & gStr

' ｼｽﾃﾑ名
    Call ItemsClearIRN
    IRN.bk = "System": IRN.rno = 0
    If IRNRead("", 1) Then
       IRN.oyian = Trim(iB1_SysName)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(iB1_SysName)
       Call IRNInsert
    End If
' ﾛｺﾞ表示
    gStr = "N" & Trim(iB1_Logo)
    If Ch_LogoDisp.Value = 1 Then
        gStr = "Y" & Trim(iB1_Logo)
    End If
' ﾕｰｻﾞｰｺｰﾄﾞ
    Call ItemsClearIRN
    IRN.bk = "Company": IRN.rno = 0
    If IRNRead("", 1) Then
       IRN.oyian = Trim(iB1_UserID)
       IRN.oyian2 = gStr
       Call IRNUpdate
    Else
       IRN.oyian = Trim(iB1_UserID)
       IRN.oyian2 = gStr
       Call IRNInsert
    End If
' ﾕｰｻﾞｰ名
    IRN.bk = "Company": IRN.rno = 1
    If IRNRead("", 1) Then
       IRN.oyian = Trim(iB1_UserName)
       Call IRNUpdate
    Else
       IRN.oyian = Trim(iB1_UserName)
       Call IRNInsert
    End If
    
    With CisFun
        .INI_FileID = gIniExe & gIniName
        .INI_Section = "DataBase"
        
        .INI_Key = "Server"
        .INI_String = iB1_Server
        .PutIni
        
        .INI_Key = "Name"
        .INI_String = iB1_DBName
        .PutIni
        
        .INI_Key = "User"
        .INI_String = iB1_UID
        .PutIni
'*--------------------------------------------------<< 2004/09/24 >> Insert Start
'   その他情報
        .INI_Section = "Others"
        ' ロゴ表示
        .INI_Key = "LogoDisp"
        Select Case Ch_LogoDisp.Value
            Case 1:    .INI_String = "Yes"
            Case Else: .INI_String = "No"
        End Select
        .PutIni
        ' ロゴ
        .INI_Key = "Logo"
        .INI_String = iB1_Logo
        .PutIni
'*--------------------------------------------------<< 2004/09/24 >> Insert End
    End With
        
    gStr = CisDB.PassWord(RTrim$(iB1_Pass))
    Call SaveSetting(gRegKey, "Environment", "DrowsSap", gStr)

' 更新確認
    MsgBox " " & vbCrLf & "         環境設定更新終了 ！！        " & vbCrLf & " ", _
            vbOKOnly + vbInformation, "環境設定更"
    B1_Area1.Enabled = True
    iB1_SysName.SetFocus
    
    On Error GoTo 0
    
    Call CisDB.DBDISConnect
    
'    sPassSkip = True
    '+----------------------------------------------------------------- 2006/03/14 Insert
    CisDB.PassLink
    '+----------------------------------------------------------------- 2006/03/14 Insert End
    If IniGet Then
       DBPut = True
       If gEnvFlg = "Err" Then
          gEnvFlg = "OK"
       End If
    End If
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
''''+----------------------+
''''+    パスワード変更     +
''''+----------------------+
'''Private Sub cmdPass_Click()
'''    frmssap.Show vbModal
'''    Unload frmssap
'''    Set frmssap = Nothing
'''End Sub
