VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CTOP0000 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "SIMPR-A Topss 環境設定"
   ClientHeight    =   5670
   ClientLeft      =   5490
   ClientTop       =   3375
   ClientWidth     =   8010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   8010
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7545
      Top             =   165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   6375
      ScaleHeight     =   555
      ScaleWidth      =   1515
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   15
      Width           =   1575
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   555
         Left            =   750
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
         Left            =   0
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
   End
   Begin VB.PictureBox B1_Area2 
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   7920
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   675
      Width           =   7980
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   2235
         Index           =   0
         Left            =   15
         Top             =   420
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   3942
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   885
            Left            =   30
            Top             =   1320
            Width           =   5745
            _ExtentX        =   10134
            _ExtentY        =   1561
            BackColor       =   14737632
            Caption         =   "退避ファイル格納先(SIMPR-A)"
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
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   100
            Begin CisText_V60.CisText iH1_Dir 
               Height          =   345
               Index           =   1
               Left            =   60
               TabIndex        =   1
               Top             =   405
               Width           =   5295
               _ExtentX        =   9340
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
               Text            =   $"CTOP0000.frx":0000
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_Dir 
               Height          =   315
               Index           =   1
               Left            =   5280
               Top             =   390
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
               cButton         =   -1  'True
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   885
            Index           =   0
            Left            =   5775
            Top             =   1320
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   1561
            BackColor       =   14737632
            Caption         =   "退避データ保持期間"
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
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   100
            Begin CisText_V60.CisText iH1_Kikan 
               Height          =   345
               Index           =   0
               Left            =   855
               TabIndex        =   2
               Top             =   450
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   609
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cCaps           =   2
               cFdAutoFormat   =   1
               cGFormat        =   "###"
               cILength        =   3
            End
            Begin VB.Label Label1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "日"
               Height          =   180
               Left            =   1380
               TabIndex        =   10
               Top             =   540
               Width           =   180
            End
         End
         Begin Cis3D_v60.CIS3D PFXLabel1 
            Height          =   405
            Index           =   2
            Left            =   -15
            Top             =   0
            Width           =   7890
            _ExtentX        =   13917
            _ExtentY        =   714
            BackColor       =   4210752
            ForeColor       =   16777215
            Caption         =   "《   受 注   》"
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
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   885
            Left            =   30
            Top             =   420
            Width           =   7830
            _ExtentX        =   13811
            _ExtentY        =   1561
            BackColor       =   14737632
            Caption         =   "連携ファイル格納先(Topss)"
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
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   100
            Begin CisText_V60.CisText iH1_Dir 
               Height          =   345
               Index           =   0
               Left            =   60
               TabIndex        =   0
               Top             =   405
               Width           =   7425
               _ExtentX        =   13097
               _ExtentY        =   609
               cWildeCard      =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   $"CTOP0000.frx":00FE
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_Dir 
               Height          =   315
               Index           =   0
               Left            =   7350
               Top             =   405
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
               cButton         =   -1  'True
            End
         End
      End
      Begin Cis3D_v60.CIS3D PFXLabel1 
         Height          =   420
         Index           =   4
         Left            =   0
         Top             =   0
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   741
         BackColor       =   16711680
         ForeColor       =   16777215
         Caption         =   "【 クライアント設定 】"
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
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   2235
         Index           =   1
         Left            =   15
         Top             =   2670
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   3942
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   885
            Left            =   30
            Top             =   1320
            Width           =   5745
            _ExtentX        =   10134
            _ExtentY        =   1561
            BackColor       =   14737632
            Caption         =   "退避ファイル格納先(SIMPR-A)"
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
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   100
            Begin CisText_V60.CisText iH1_Dir 
               Height          =   345
               Index           =   3
               Left            =   60
               TabIndex        =   4
               Top             =   405
               Width           =   5295
               _ExtentX        =   9340
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
               Text            =   $"CTOP0000.frx":01FC
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_Dir 
               Height          =   315
               Index           =   3
               Left            =   5280
               Top             =   390
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
               cButton         =   -1  'True
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   885
            Index           =   1
            Left            =   5775
            Top             =   1320
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   1561
            BackColor       =   14737632
            Caption         =   "退避データ保持期間"
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
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   100
            Begin CisText_V60.CisText iH1_Kikan 
               Height          =   345
               Index           =   1
               Left            =   855
               TabIndex        =   5
               Top             =   450
               Width           =   465
               _ExtentX        =   820
               _ExtentY        =   609
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cCaps           =   2
               cFdAutoFormat   =   1
               cGFormat        =   "###"
               cILength        =   3
            End
            Begin VB.Label Label2 
               BackColor       =   &H00E0E0E0&
               Caption         =   "日"
               Height          =   180
               Left            =   1380
               TabIndex        =   11
               Top             =   540
               Width           =   180
            End
         End
         Begin Cis3D_v60.CIS3D PFXLabel1 
            Height          =   405
            Index           =   1
            Left            =   -15
            Top             =   0
            Width           =   7890
            _ExtentX        =   13917
            _ExtentY        =   714
            BackColor       =   4210752
            ForeColor       =   16777215
            Caption         =   "《   出 荷   》"
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
         Begin Cis3D_v60.CIS3D CIS3D8 
            Height          =   885
            Left            =   30
            Top             =   420
            Width           =   7830
            _ExtentX        =   13811
            _ExtentY        =   1561
            BackColor       =   14737632
            Caption         =   "連携ファイル格納先(Topps)"
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
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   100
            Begin CisText_V60.CisText iH1_Dir 
               Height          =   345
               Index           =   2
               Left            =   60
               TabIndex        =   3
               Top             =   405
               Width           =   7425
               _ExtentX        =   13097
               _ExtentY        =   609
               cWildeCard      =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   $"CTOP0000.frx":02FA
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_Dir 
               Height          =   315
               Index           =   2
               Left            =   7350
               Top             =   405
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
               cButton         =   -1  'True
            End
         End
      End
   End
   Begin Cis3D_v60.CIS3D CIS3D2 
      Height          =   675
      Left            =   0
      Top             =   0
      Width           =   7980
      _ExtentX        =   14076
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
      Begin VB.Timer Timer1 
         Left            =   0
         Top             =   480
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   495
         Left            =   30
         Top             =   90
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   873
         ForeColor       =   0
         Caption         =   "Topps連携 環境設定"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "HGSｺﾞｼｯｸE"
            Size            =   18
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "SIMPR-A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   825
         Left            =   3495
         TabIndex        =   7
         Top             =   -30
         Width           =   2730
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "SIMPR-A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   825
         Left            =   3420
         TabIndex        =   6
         Top             =   -90
         Width           =   2730
      End
   End
End
Attribute VB_Name = "CTOP0000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  TOPPS連携　環境設定
'**       フォームID    :  CTOP0000
'**
'**       処理概要      :
'**
'**       作  成  日    :  2004/04/15  By CIS
'**       変  更  日    :  2007/01/12  ナルコ対応
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ

    Dim mGotColor       As Long
    Dim mLostColor      As Long
    Dim Rg_On           As String
    Dim Rg_Off          As String
    Dim wFlg            As Boolean
    Dim wI              As Byte
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

'   #------------------#
'   # 初 期 内 容 取 得 #
'   #------------------#
    If Not gRtn Then
        If Not IniGet Then gEnvFlg = "Err"
    End If

'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Call CisFrm.InitFld

'   #------------------#
'   #   画面中央表示    #
'   #------------------#
    Call CisFrm.Frm_Center(False)
    
    mGotColor = &HC0FFC0
    mLostColor = &HE0E0E0
'   ( 内容表示 )
    Call BodyDisp

'   ( 組込起動 )
    If Not gRtn Then Me.Show Else PB_END.BtnCaption = "戻る"
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
'   ( タイトルバー用 )
    With CisFrm
        .TimeSet_TitleName = "SIMPR-A TOPPS連携 環境設定"
        .TimeSet_TitleSpace = 100
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
        Case vbKeyF12:      If PB_ENT.Visible Then Call PB_END_Click    '【終了】
        Case vbKeyReturn:   If PB_ENT.Visible Then Call PB_ENT_Click    '【入力】
        Case Else:          Exit Function
    End Select
    
Key_Acc_ED:
    Key_Acc = True
End Function

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
    If Not gRtn Then
        Call CisDB.DBDISConnect
        Call CisFrm.UnLoadDisp
        End
    Else
        Me.Hide
    End If
End Sub

'+-------------------------------+
'+   ファイル出力先　参照　+
'+-------------------------------+
Private Sub PB_Dir_Click(Index As Integer)
    With CisFun
        Select Case Index
            Case 0: .Folder_Title = "【受注】連携ファイル格納先(Topss)参照"
            Case 1: .Folder_Title = "【受注】退避ﾌｧｲﾙ格納先(SIMPR-A)参照"
            Case 2: .Folder_Title = "【出荷】連携ファイル格納先(Topss)参照"
            Case 3: .Folder_Title = "【出荷】退避ﾌｧｲﾙ格納先(SIMPR-A)参照"
        End Select
        .Folder_DefDir = iH1_Dir(Index)
        .FolderSel
        
        If RTrim$(.Folder_Result) <> "" Then iH1_Dir(Index) = RTrim$(.Folder_Result)
    End With
    iH1_Dir(Index).SetFocus
End Sub
'****************************
'*   出荷データベース名参照
'****************************
'Private Sub PB_DBName_Click()
'     On Error Resume Next
'     CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly
'     CommonDialog1.FileName = gIniAccessName
'     CommonDialog1.Filter = "Microsoft Access ﾃﾞｰﾀﾍﾞｰｽ (*.mdb)| *.mdb"
'    ' 以下のｿｰｽｺｰﾄﾞは、<ｷｬﾝｾﾙ>ﾎﾞﾀﾝがｸﾘｯｸされたときに実行されます。
'     CommonDialog1.CancelError = True
'     CommonDialog1.ShowOpen
'     If ERR = 32755 Then
'    '<ｷｬﾝｾﾙ>ﾎﾞﾀﾝが押されたときにｴﾗｰｺｰﾄﾞを返すようにする設定です。
'    '必要なｿｰｽｺｰﾄﾞを入力してください。
'     End If
'     If Trim(CommonDialog1.FileName) <> "" Then
'        gIniAccessName = CommonDialog1.FileName
'        iH1_Dir(2) = gIniAccessName
'     End If
'     On Error GoTo 0
'End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If Not BodyChk Then GoTo ReturnPress_Ed
    If Not DBPut Then GoTo ReturnPress_Ed
ReturnPress_Ed:
    Call CisFrm.MousePT(1)
End Sub
'****************************
'*      BODYﾁｪｯｸ処理
'****************************
Private Function BodyChk() As Boolean
    BodyChk = False
'(( 受注 ))
    For gInt = 0 To 3
        If Trim(iH1_Dir(gInt)) = "" Then
           Select Case gInt
                Case 0
                    gStr = "　(( 受注 )) 連携ファイル格納先(Topss)を入力して下さい。　"
                Case 1
                    gStr = "　(( 受注 )) 退避ファイル格納先(SIMPR-A)を入力して下さい。　"
                Case 2
                    gStr = "　(( 出荷 )) 連携ファイル格納先(Topss)を入力して下さい。　"
                Case 3
                    gStr = "　(( 出荷 )) 退避ファイル格納先(SIMPR-A)を入力して下さい。　"
           End Select
                    
           MsgBox gStr, vbExclamation, "必須入力ｴﾗｰ"
           iH1_Dir(gInt).SetFocus
           Exit Function
        End If
        
        ' ディレクトリ存在確認
        With CisFun
'            If gInt = 0 Or gInt = 2 Then
'                If Right(iH1_Dir(gInt), 1) <> "\" Then
'                    .FE_FullPath = RTrim(iH1_Dir(gInt))
'                Else
'                    .FE_FullPath = RTrim(iH1_Dir(gInt))
'               End If
''               .GetPathName
'            Else
            If Right(iH1_Dir(gInt), 1) <> "\" Then
                .FE_Path = RTrim(iH1_Dir(gInt)) & "\"
            Else
                .FE_Path = RTrim(iH1_Dir(gInt))
            End If
'            End If
            If gInt = 0 Or gInt = 2 Then
               .DirCheck NoMsg_NoCrt
            Else
               .DirCheck Msg_Crt
            End If
            If Not .DirCheck Then
               Select Case gInt
                     Case 0
                         gStr = "　(( 受注 )) 存在する連携ファイル格納先(Topss)を入力して下さい。　"
                     Case 1
                         gStr = "　(( 受注 )) 存在する退避ファイル格納先(SIMPR-A)を入力して下さい。　"
                     Case 2
                         gStr = "　(( 出荷 )) 存在する連携ファイル格納先(Topss)を入力して下さい。　"
                     Case 3
                         gStr = "　(( 出荷 )) 存在する退避ファイル格納先(SIMPR-A)を入力して下さい。　"
               End Select
               MsgBox gStr, vbExclamation, "必須入力ｴﾗｰ"
               iH1_Dir(gInt).SetFocus
               Exit Function
            End If
        End With
    Next gInt
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
'
    Call ToppsIniGet("SET")
    
    ' 受注
    iH1_Dir(0) = TJyu_Dir
    iH1_Dir(1) = TJyu_SaveDir
    iH1_Kikan(0) = CisFun.RSetFld(TJyu_SaveDel, 0, iH1_Kikan(0).cFormat)

    ' 出荷
    iH1_Dir(2) = TSyu_Dir
    iH1_Dir(3) = TSyu_SaveDir
    iH1_Kikan(1) = CisFun.RSetFld(TSyu_SaveDel, 0, iH1_Kikan(1).cFormat)
End Sub
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()

    DBPut = False
'  ( 確認メッセージ )
    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = "     【 TOPPS連携 設定変更確認 】      "
        .MB_MSG(3) = "    TOPPS連携 設定を変更しますか      "
        .MB_Title = "TOPPS連携設定"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With

    With CisFun
        ' 受注
        .INI_Section = "Topps"
        .INI_FileID = gIniExe & gIniName
        
        .INI_Key = "JyuDir"
        .INI_String = RTrim(iH1_Dir(0))
        .PutIni
        .INI_Key = "JyuSaveDir"
        .INI_String = RTrim(iH1_Dir(1))
        .PutIni
        .INI_Key = "JyuSaveDelDay"
        .INI_String = Format(iH1_Kikan(0), "0")
        .PutIni
        ' 出荷
        .INI_Key = "SyuDir"
        .INI_String = RTrim(iH1_Dir(2))
        .PutIni
        .INI_Key = "SyuSaveDir"
        .INI_String = RTrim(iH1_Dir(3))
        .PutIni
        .INI_Key = "SyuSaveDelDay"
        .INI_String = Format(iH1_Kikan(1), "0")
        .PutIni
    End With
    
        
' 更新確認
    MsgBox " " & vbCrLf & "         TOPPS連携 設定更新終了 ！！        " & vbCrLf & " ", _
            vbOKOnly + vbInformation, "TOPPS連携 設定変更"
'    B1_Area1.Enabled = True
    
    On Error GoTo 0
    
'    Call CisDB.DBDISConnect
    
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function

