VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CXJ0700 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "SIMPR-A テキスト連携 環境設定"
   ClientHeight    =   7500
   ClientLeft      =   2205
   ClientTop       =   1650
   ClientWidth     =   15870
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   15870
   Begin VB.PictureBox B1_Area2 
      Height          =   6630
      Left            =   0
      ScaleHeight     =   6570
      ScaleWidth      =   15825
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   690
      Width           =   15885
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   3420
         Index           =   0
         Left            =   15
         Top             =   15
         Width           =   15795
         _ExtentX        =   27861
         _ExtentY        =   6033
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
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   300
            Index           =   0
            Left            =   0
            Top             =   405
            Width           =   7905
            _ExtentX        =   13944
            _ExtentY        =   529
            BackColor       =   16761087
            Caption         =   "※　受注テキスト環境"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cPositionX      =   60
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   885
            Index           =   2
            Left            =   30
            Top             =   2505
            Width           =   5880
            _ExtentX        =   10372
            _ExtentY        =   1561
            BackColor       =   14737632
            Caption         =   "変換リスト発行"
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
            cBoderStyle     =   1
            cPositionX      =   100
            cPositionY      =   100
            Begin VB.Frame Back_Prt 
               BackColor       =   &H00E0E0E0&
               Height          =   555
               Index           =   0
               Left            =   1170
               TabIndex        =   40
               Top             =   300
               Width           =   4665
               Begin VB.CheckBox H1Ch_Err 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "エラー内容"
                  Height          =   345
                  Index           =   0
                  Left            =   90
                  TabIndex        =   5
                  Top             =   150
                  Width           =   1275
               End
               Begin VB.CheckBox H1Ch_Kari 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "仮登録内容"
                  Height          =   345
                  Index           =   0
                  Left            =   1380
                  TabIndex        =   6
                  Top             =   150
                  Width           =   1275
               End
               Begin VB.CheckBox H1Ch_Etc 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "その他"
                  Height          =   345
                  Index           =   0
                  Left            =   3690
                  TabIndex        =   8
                  Top             =   150
                  Width           =   855
               End
               Begin VB.CheckBox H1Ch_Dbl 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "重複"
                  Height          =   345
                  Index           =   0
                  Left            =   2820
                  TabIndex        =   7
                  Top             =   150
                  Width           =   855
               End
            End
            Begin VB.CheckBox H1ch_Prt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "なし"
               Height          =   345
               Index           =   0
               Left            =   270
               TabIndex        =   4
               Top             =   420
               Width           =   705
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   885
            Index           =   7
            Left            =   6030
            Top             =   705
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   1561
            BackColor       =   14737632
            Caption         =   "連携ﾌｧｲﾙ名"
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
            Begin CisText_V60.CisText iH1_TopFName 
               Height          =   345
               Index           =   0
               Left            =   120
               TabIndex        =   2
               Top             =   420
               Width           =   1635
               _ExtentX        =   2884
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
               Text            =   "XXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   20
               cDataReplace    =   1
               cCaps           =   2
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   885
            Index           =   0
            Left            =   30
            Top             =   1605
            Width           =   7845
            _ExtentX        =   13838
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
            Begin CisText_V60.CisText iH1_BackDir 
               Height          =   345
               Index           =   0
               Left            =   60
               TabIndex        =   3
               Top             =   405
               Width           =   7485
               _ExtentX        =   13203
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
               Text            =   $"CXJ0700.frx":0000
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_BackDir 
               Height          =   315
               Index           =   0
               Left            =   7440
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
            Left            =   5910
            Top             =   2505
            Width           =   1980
            _ExtentX        =   3493
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
               Left            =   735
               TabIndex        =   9
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
               Left            =   1260
               TabIndex        =   41
               Top             =   540
               Width           =   180
            End
         End
         Begin Cis3D_v60.CIS3D PFXLabel1 
            Height          =   405
            Index           =   2
            Left            =   -15
            Top             =   0
            Width           =   15765
            _ExtentX        =   27808
            _ExtentY        =   714
            BackColor       =   4210752
            ForeColor       =   16777215
            Caption         =   "《   【愛知機械】 テキスト取り込み 》"
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
            Index           =   0
            Left            =   30
            Top             =   705
            Width           =   6000
            _ExtentX        =   10583
            _ExtentY        =   1561
            BackColor       =   14737632
            Caption         =   "連携ファイル格納先"
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
               TabIndex        =   1
               Top             =   405
               Width           =   5655
               _ExtentX        =   9975
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
               Text            =   $"CXJ0700.frx":00FE
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_Dir 
               Height          =   315
               Index           =   0
               Left            =   5610
               Top             =   420
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
            Index           =   5
            Left            =   13905
            Top             =   705
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   1561
            BackColor       =   14737632
            Caption         =   "連携ﾌｧｲﾙ名"
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
            Begin CisText_V60.CisText iH1_TopFName 
               Height          =   345
               Index           =   2
               Left            =   120
               TabIndex        =   11
               Top             =   420
               Width           =   1635
               _ExtentX        =   2884
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
               Text            =   "XXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   20
               cDataReplace    =   1
               cCaps           =   2
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   885
            Index           =   1
            Left            =   7905
            Top             =   705
            Width           =   6000
            _ExtentX        =   10583
            _ExtentY        =   1561
            BackColor       =   14737632
            Caption         =   "連携ファイル格納先"
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
               TabIndex        =   10
               Top             =   405
               Width           =   5655
               _ExtentX        =   9975
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
               Text            =   $"CXJ0700.frx":01FC
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_Dir 
               Height          =   315
               Index           =   2
               Left            =   5610
               Top             =   420
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
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   300
            Index           =   1
            Left            =   7905
            Top             =   405
            Width           =   7860
            _ExtentX        =   13864
            _ExtentY        =   529
            BackColor       =   16761024
            Caption         =   "※　出荷テキスト環境"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cPositionX      =   60
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   885
            Index           =   1
            Left            =   7905
            Top             =   1605
            Width           =   7845
            _ExtentX        =   13838
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
            Begin CisText_V60.CisText iH1_BackDir 
               Height          =   345
               Index           =   2
               Left            =   60
               TabIndex        =   12
               Top             =   405
               Width           =   7485
               _ExtentX        =   13203
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
               Text            =   $"CXJ0700.frx":02FA
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_BackDir 
               Height          =   315
               Index           =   2
               Left            =   7440
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
            Index           =   6
            Left            =   7905
            Top             =   2505
            Width           =   5880
            _ExtentX        =   10372
            _ExtentY        =   1561
            BackColor       =   14737632
            Caption         =   "変換リスト発行"
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
            cBoderStyle     =   1
            cPositionX      =   100
            cPositionY      =   100
            Begin VB.CheckBox H1ch_Prt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "なし"
               Height          =   345
               Index           =   2
               Left            =   270
               TabIndex        =   13
               Top             =   420
               Width           =   705
            End
            Begin VB.Frame Back_Prt 
               BackColor       =   &H00E0E0E0&
               Height          =   555
               Index           =   2
               Left            =   1170
               TabIndex        =   44
               Top             =   300
               Width           =   4665
               Begin VB.CheckBox H1Ch_Dbl 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "重複"
                  Height          =   345
                  Index           =   2
                  Left            =   2820
                  TabIndex        =   16
                  Top             =   150
                  Width           =   855
               End
               Begin VB.CheckBox H1Ch_Etc 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "その他"
                  Height          =   345
                  Index           =   2
                  Left            =   3690
                  TabIndex        =   17
                  Top             =   150
                  Width           =   855
               End
               Begin VB.CheckBox H1Ch_Kari 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "仮登録内容"
                  Height          =   345
                  Index           =   2
                  Left            =   1380
                  TabIndex        =   15
                  Top             =   150
                  Width           =   1275
               End
               Begin VB.CheckBox H1Ch_Err 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "エラー内容"
                  Height          =   345
                  Index           =   2
                  Left            =   90
                  TabIndex        =   14
                  Top             =   150
                  Width           =   1275
               End
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   885
            Index           =   8
            Left            =   13785
            Top             =   2505
            Width           =   1980
            _ExtentX        =   3493
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
               Index           =   2
               Left            =   735
               TabIndex        =   18
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
            Begin VB.Label Label5 
               BackColor       =   &H00E0E0E0&
               Caption         =   "日"
               Height          =   180
               Left            =   1260
               TabIndex        =   45
               Top             =   540
               Width           =   180
            End
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   3135
         Index           =   1
         Left            =   15
         Top             =   3435
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   5530
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
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   885
            Index           =   1
            Left            =   30
            Top             =   2220
            Width           =   5880
            _ExtentX        =   10372
            _ExtentY        =   1561
            BackColor       =   14737632
            Caption         =   "変換リスト発行"
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
            cBoderStyle     =   1
            cPositionX      =   100
            cPositionY      =   100
            Begin VB.CheckBox H1ch_Prt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "なし"
               Height          =   345
               Index           =   1
               Left            =   300
               TabIndex        =   22
               Top             =   420
               Width           =   705
            End
            Begin VB.Frame Back_Prt 
               BackColor       =   &H00E0E0E0&
               Height          =   555
               Index           =   1
               Left            =   1170
               TabIndex        =   42
               Top             =   300
               Width           =   4665
               Begin VB.CheckBox H1Ch_Dbl 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "重複"
                  Height          =   345
                  Index           =   1
                  Left            =   2820
                  TabIndex        =   25
                  Top             =   150
                  Width           =   855
               End
               Begin VB.CheckBox H1Ch_Etc 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "その他"
                  Height          =   345
                  Index           =   1
                  Left            =   3690
                  TabIndex        =   26
                  Top             =   150
                  Width           =   855
               End
               Begin VB.CheckBox H1Ch_Kari 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "仮登録内容"
                  Height          =   345
                  Index           =   1
                  Left            =   1380
                  TabIndex        =   24
                  Top             =   150
                  Width           =   1275
               End
               Begin VB.CheckBox H1Ch_Err 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "エラー内容"
                  Height          =   345
                  Index           =   1
                  Left            =   90
                  TabIndex        =   23
                  Top             =   150
                  Width           =   1275
               End
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   885
            Index           =   3
            Left            =   6030
            Top             =   420
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   1561
            BackColor       =   14737632
            Caption         =   "連携ﾌｧｲﾙ名"
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
            Begin CisText_V60.CisText iH1_TopFName 
               Height          =   345
               Index           =   1
               Left            =   120
               TabIndex        =   20
               Top             =   420
               Width           =   1635
               _ExtentX        =   2884
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
               Text            =   "XXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   20
               cDataReplace    =   1
               cCaps           =   2
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   885
            Left            =   30
            Top             =   1320
            Width           =   7845
            _ExtentX        =   13838
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
            Begin CisText_V60.CisText iH1_BackDir 
               Height          =   345
               Index           =   1
               Left            =   60
               TabIndex        =   21
               Top             =   405
               Width           =   7485
               _ExtentX        =   13203
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
               Text            =   $"CXJ0700.frx":03F8
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_BackDir 
               Height          =   315
               Index           =   1
               Left            =   7440
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
            Index           =   4
            Left            =   5910
            Top             =   2220
            Width           =   1980
            _ExtentX        =   3493
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
               Left            =   735
               TabIndex        =   27
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
               Left            =   1260
               TabIndex        =   43
               Top             =   540
               Width           =   180
            End
         End
         Begin Cis3D_v60.CIS3D PFXLabel1 
            Height          =   405
            Index           =   0
            Left            =   -15
            Top             =   0
            Width           =   7890
            _ExtentX        =   13917
            _ExtentY        =   714
            BackColor       =   4210752
            ForeColor       =   16777215
            Caption         =   "《   【テクノエイト】 テキスト取り込み   》"
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
            Width           =   6000
            _ExtentX        =   10583
            _ExtentY        =   1561
            BackColor       =   14737632
            Caption         =   "連携ファイル格納先"
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
               TabIndex        =   19
               Top             =   405
               Width           =   5655
               _ExtentX        =   9975
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
               Text            =   $"CXJ0700.frx":04F6
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_Dir 
               Height          =   315
               Index           =   1
               Left            =   5610
               Top             =   420
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
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   3135
         Index           =   2
         Left            =   7890
         Top             =   3435
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   5530
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
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   885
            Index           =   9
            Left            =   30
            Top             =   2220
            Width           =   5880
            _ExtentX        =   10372
            _ExtentY        =   1561
            BackColor       =   14737632
            Caption         =   "変換リスト発行"
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
            cBoderStyle     =   1
            cPositionX      =   100
            cPositionY      =   100
            Begin VB.Frame Back_Prt 
               BackColor       =   &H00E0E0E0&
               Height          =   555
               Index           =   3
               Left            =   1170
               TabIndex        =   46
               Top             =   300
               Width           =   4665
               Begin VB.CheckBox H1Ch_Err 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "エラー内容"
                  Height          =   345
                  Index           =   3
                  Left            =   90
                  TabIndex        =   32
                  Top             =   150
                  Width           =   1275
               End
               Begin VB.CheckBox H1Ch_Kari 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "仮登録内容"
                  Height          =   345
                  Index           =   3
                  Left            =   1380
                  TabIndex        =   33
                  Top             =   150
                  Width           =   1275
               End
               Begin VB.CheckBox H1Ch_Etc 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "その他"
                  Height          =   345
                  Index           =   3
                  Left            =   3690
                  TabIndex        =   35
                  Top             =   150
                  Width           =   855
               End
               Begin VB.CheckBox H1Ch_Dbl 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "重複"
                  Height          =   345
                  Index           =   3
                  Left            =   2820
                  TabIndex        =   34
                  Top             =   150
                  Width           =   855
               End
            End
            Begin VB.CheckBox H1ch_Prt 
               BackColor       =   &H00E0E0E0&
               Caption         =   "なし"
               Height          =   345
               Index           =   3
               Left            =   300
               TabIndex        =   31
               Top             =   420
               Width           =   705
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   885
            Index           =   10
            Left            =   6030
            Top             =   420
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   1561
            BackColor       =   14737632
            Caption         =   "連携ﾌｧｲﾙ名"
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
            Begin CisText_V60.CisText iH1_TopFName 
               Height          =   345
               Index           =   3
               Left            =   120
               TabIndex        =   29
               Top             =   420
               Width           =   1635
               _ExtentX        =   2884
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
               Text            =   "XXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   20
               cDataReplace    =   1
               cCaps           =   2
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   885
            Left            =   30
            Top             =   1320
            Width           =   7845
            _ExtentX        =   13838
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
            Begin CisText_V60.CisText iH1_BackDir 
               Height          =   345
               Index           =   3
               Left            =   60
               TabIndex        =   30
               Top             =   405
               Width           =   7485
               _ExtentX        =   13203
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
               Text            =   $"CXJ0700.frx":05F4
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_BackDir 
               Height          =   315
               Index           =   3
               Left            =   7440
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
            Index           =   11
            Left            =   5910
            Top             =   2220
            Width           =   1980
            _ExtentX        =   3493
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
               Index           =   3
               Left            =   735
               TabIndex        =   36
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
            Begin VB.Label Label6 
               BackColor       =   &H00E0E0E0&
               Caption         =   "日"
               Height          =   180
               Left            =   1260
               TabIndex        =   47
               Top             =   540
               Width           =   180
            End
         End
         Begin Cis3D_v60.CIS3D PFXLabel1 
            Height          =   405
            Index           =   1
            Left            =   15
            Top             =   0
            Width           =   7875
            _ExtentX        =   13891
            _ExtentY        =   714
            BackColor       =   4210752
            ForeColor       =   16777215
            Caption         =   "《   【ヨロズ愛知】 テキスト取り込み   》"
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
         Begin Cis3D_v60.CIS3D CIS3D11 
            Height          =   885
            Left            =   30
            Top             =   420
            Width           =   6000
            _ExtentX        =   10583
            _ExtentY        =   1561
            BackColor       =   14737632
            Caption         =   "連携ファイル格納先"
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
               TabIndex        =   28
               Top             =   405
               Width           =   5655
               _ExtentX        =   9975
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
               Text            =   $"CXJ0700.frx":06F2
               MaxLength       =   250
               IMEMode         =   2
               cCaps           =   2
            End
            Begin Cis3D_v60.CIS3D PB_Dir 
               Height          =   315
               Index           =   3
               Left            =   5610
               Top             =   420
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
      Left            =   14220
      ScaleHeight     =   555
      ScaleWidth      =   1515
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   30
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
   Begin Cis3D_v60.CIS3D CIS3D2 
      Height          =   675
      Left            =   0
      Top             =   0
      Width           =   15840
      _ExtentX        =   27940
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
         Left            =   135
         Top             =   90
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   873
         ForeColor       =   0
         Caption         =   "テキスト連携 環境設定"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "HGSｺﾞｼｯｸE"
            Size            =   15.75
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
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   4500
         TabIndex        =   37
         Top             =   105
         Width           =   1680
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "SIMPR-A"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   510
         Left            =   4470
         TabIndex        =   0
         Top             =   75
         Width           =   1680
      End
   End
End
Attribute VB_Name = "CXJ0700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  受注テキスト連携　環境設定
'**       フォームID    :  PXJ0700  (Based PTP3000)
'**
'**       処理概要      :
'**
'**       作  成  日    :  2009/09/14  By CIS
'**       変  更  日    :  2010/03/17  By CIS  愛知機械：出荷環境追加
'**       変  更  日    :  2010/04/19  By CIS  ヨロズ愛知：受注環境追加
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

    Dim eKanban_Proc           As Byte     ' 0:受注(愛知機械),1:受注(ﾃｸﾉｴｲﾄ),2:出荷(愛知機械)
'    Dim eKanban_Prefix(1)      As String   ' ﾌﾟﾚﾌｨｯｸｽ (0:eJY,1:eSK)
'    Dim eKanban_PrefixNm(1)    As String   ' ﾌﾟﾚﾌｨｯｸｽ名 (0:受注,1:出荷)
'    Dim eKanban_Orderer(1)     As String   ' 受注先(0:愛知機械,1:テクノエイト)  ' 2009.09.14 add
'    Dim eKanban_Prefix(2)      As String   ' ﾌﾟﾚﾌｨｯｸｽ (0:eJY(愛知機械),1:eJY(ﾃｸﾉｴｲﾄ),2:eSK(愛知機械)
'    Dim eKanban_PrefixNm(2)    As String   ' ﾌﾟﾚﾌｨｯｸｽ名 (0:受注(愛知機械),1:受注(ﾃｸﾉｴｲﾄ),2:出荷(愛知機械)
'    Dim eKanban_Orderer(2)     As String   ' 受注先(0:愛知機械,1:テクノエイト,2:愛知機械)
    
    Dim eKanban_Prefix(3)      As String   ' ﾌﾟﾚﾌｨｯｸｽ (0:eJY(愛知機械),1:eJY(ﾃｸﾉｴｲﾄ),2:eSK(愛知機械)
    Dim eKanban_PrefixNm(3)    As String   ' ﾌﾟﾚﾌｨｯｸｽ名 (0:受注(愛知機械),1:受注(ﾃｸﾉｴｲﾄ),2:出荷(愛知機械)
    Dim eKanban_Orderer(3)     As String   ' 受注先(0:愛知機械,1:テクノエイト,2:愛知機械)
    
    Dim mBlen           As Boolean
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   # 二 重 起 動 防 止 #
'   #-------------------#
    If Not Dupli_Start Then End

    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me

    Timer1.Interval = 1000          ' 時刻表示間隔
    Timer1.Enabled = False          ' 時刻表示開始  ' 2009.09.15 upd : True-->False (あとから消しているのでFalseとした)

'   #-------------------#
'   # 初 期 内 容 取 得 #
'   #-------------------#
    If Not gRtn Then
        If Not IniGet Then gEnvFlg = "Err"
    End If

'   #-------------------#
'   #  項 目 初 期 化   #
'   #-------------------#
    Call CisFrm.InitFld

'   #-------------------#
'   #   画面中央表示    #
'   #-------------------#
    Call CisFrm.Frm_Center(False)
    
    mGotColor = &HC0FFC0
    mLostColor = &HE0E0E0
    
    ' 処理名をセット
    eKanban_Prefix(0) = "eJY": eKanban_PrefixNm(0) = "受注"
    eKanban_Prefix(1) = "eJY": eKanban_PrefixNm(1) = "受注"     '
    eKanban_Prefix(2) = "eSK": eKanban_PrefixNm(2) = "出荷"     '
    
    eKanban_Prefix(3) = "eJY": eKanban_PrefixNm(3) = "受注"         '2010/04/19
    
    
    ' 受注先をセット                                            '
    eKanban_Orderer(0) = "愛知機械"                             '
    eKanban_Orderer(1) = "テクノエイト"                         '
    eKanban_Orderer(2) = "愛知機械"                             '
    
    eKanban_Orderer(3) = "ヨロズ愛知"                             ' 2010/04/19
    
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
        .TimeSet_TitleName = "SIMPR-A テキスト連携 環境設定"    '
        .TimeSet_TitleSpace = 100                                   '
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

Private Sub H1ch_Prt_Click(Index As Integer)
    
    If H1ch_Prt(Index).Value = 0 Then
       H1ch_Prt(Index).Caption = "なし"
       mBlen = False
    Else
       H1ch_Prt(Index).Caption = "あり"
       mBlen = True
    End If

    H1Ch_Err(Index).Enabled = mBlen
    H1Ch_Kari(Index).Enabled = mBlen
    H1Ch_Etc(Index).Enabled = mBlen
    H1Ch_Dbl(Index).Enabled = mBlen
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
    If Not gRtn Then
        Call CisDB.DBDISConnect
        Call CisFrm.UnLoadDisp
        End
    Else
        Me.Hide
    End If
End Sub

'+-------------------------------+
'+   ファイル出力先(受注)　参照　+
'+-------------------------------+
Private Sub PB_Dir_Click(Index As Integer)
    With CisFun
        Select Case Index
            Case 0: .Folder_Title = "【" & eKanban_Orderer(0) & "】連携ファイル格納先参照"
            Case 1: .Folder_Title = "【" & eKanban_Orderer(1) & "】連携ファイル格納先参照"
            Case 2: .Folder_Title = "【" & eKanban_Orderer(2) & "】連携ファイル格納先参照"
        
            Case 3: .Folder_Title = "【" & eKanban_Orderer(3) & "】連携ファイル格納先参照"          '2010/04/19
        End Select
        .Folder_DefDir = iH1_Dir(Index)
        .FolderSel
        
        If RTrim$(.Folder_Result) <> "" Then iH1_Dir(Index) = RTrim$(.Folder_Result)
    End With
    iH1_Dir(Index).SetFocus
End Sub
'+-----------------------------------+
'+   ファイル出力先(SIMPR-A)　参照　 +
'+-----------------------------------+
Private Sub PB_BackDir_Click(Index As Integer)
    With CisFun
        Select Case Index
            Case 0: .Folder_Title = "【" & eKanban_Orderer(0) & "】退避ﾌｧｲﾙ格納先(SIMPR-A)参照"
            Case 1: .Folder_Title = "【" & eKanban_Orderer(1) & "】退避ﾌｧｲﾙ格納先(SIMPR-A)参照"
            Case 2: .Folder_Title = "【" & eKanban_Orderer(2) & "】退避ﾌｧｲﾙ格納先(SIMPR-A)参照"
            
            Case 3: .Folder_Title = "【" & eKanban_Orderer(3) & "】退避ﾌｧｲﾙ格納先(SIMPR-A)参照"     '2010/04/19
        End Select
        .Folder_DefDir = iH1_BackDir(Index)
        .FolderSel
        
        If RTrim$(.Folder_Result) <> "" Then iH1_BackDir(Index) = RTrim$(.Folder_Result)
    End With
    iH1_BackDir(Index).SetFocus
End Sub
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

'(( フォルダ確認 ))
'    For gInt = 0 To 2
    For gInt = 0 To 3                   '2010/04/19
        If Trim(iH1_Dir(gInt)) = "" Then
           gStr = "　(( " & eKanban_Orderer(gInt) & " )) 連携ファイル格納先を入力して下さい。　"
           MsgBox gStr, vbExclamation, "必須入力ｴﾗｰ"
           iH1_Dir(gInt).SetFocus
           Exit Function
        End If
        
        If Trim(iH1_TopFName(gInt)) = "" Then
           gStr = "　(( " & eKanban_Orderer(gInt) & " )) 連携ファイル名を入力して下さい。　"
           MsgBox gStr, vbExclamation, "必須入力ｴﾗｰ"
           iH1_TopFName(gInt).SetFocus
           Exit Function
        End If
        
'        If gInt <> 2 Then
'            If Trim(iH1_BackDir(gInt)) = "" Then
'               gStr = "　(( " & eKanban_Orderer(gInt) & " )) 退避ファイル格納先(SIMPR-A)を入力して下さい。　"
'               MsgBox gStr, vbExclamation, "必須入力ｴﾗｰ"
'               iH1_BackDir(gInt).SetFocus
'               Exit Function
'            End If
'        End If
        
        ' ディレクトリ存在確認
        With CisFun
            If Right(iH1_Dir(gInt), 1) <> "\" Then
                .FE_Path = RTrim(iH1_Dir(gInt)) & "\"
            Else
                .FE_Path = RTrim(iH1_Dir(gInt))
            End If
            If gInt = 0 Then
               .DirCheck NoMsg_NoCrt
            Else
               .DirCheck NoMsg_NoCrt
            End If
            If Not .DirCheck Then
               gStr = "　(( " & eKanban_Orderer(gInt) & " )) 存在する連携ファイル格納先を入力して下さい。　"
               MsgBox gStr, vbExclamation, "必須入力ｴﾗｰ"
               iH1_Dir(gInt).SetFocus
               Exit Function
            End If
            If gInt <> 2 Then
                If Right(iH1_BackDir(gInt), 1) <> "\" Then
                    .FE_Path = RTrim(iH1_BackDir(gInt)) & "\"
                Else
                    .FE_Path = RTrim(iH1_BackDir(gInt))
                End If
                .DirCheck Msg_Crt
                If Not .DirCheck Then
                   gStr = "　(( " & eKanban_Orderer(gInt) & " )) 存在する退避ファイル格納先(SIMPR-A)を入力して下さい。　"
                   MsgBox gStr, vbExclamation, "必須入力ｴﾗｰ"
                   iH1_BackDir(gInt).SetFocus
                   Exit Function
                End If
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
'    For eKanban_Proc = 0 To 2
    For eKanban_Proc = 0 To 3                               '2010/04/19
        With CisFun
            '------------------- ｸﾗｲｱﾝﾄ個別設定
            .INI_Section = "JText"
            .INI_FileID = gIniExe & gwsIniName
            '-------- 受注ﾃｷｽﾄ格納先
            .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_" & eKanban_Orderer(eKanban_Proc) & "_ﾌｧｲﾙ格納先"
            .INI_Default = ""
            .GetIni
            If Trim(.INI_String) = "" Then
               .INI_String = ""
               .PutIni
            End If
            iH1_Dir(eKanban_Proc) = Trim(.INI_String)
            
            If Trim(iH1_Dir(eKanban_Proc)) <> "" Then
               If Right(iH1_Dir(eKanban_Proc), 1) <> "\" Then
                  iH1_Dir(eKanban_Proc) = iH1_Dir(eKanban_Proc) & "\"
               End If
            End If
            '-------- 受注ﾃｷｽﾄ退避先(SIMPR-A)
            .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_" & eKanban_Orderer(eKanban_Proc) & "_ﾌｧｲﾙ退避先"
            .INI_Default = ""
            .GetIni
            If Trim(.INI_String) = "" Then
               gStr = Replace(UCase(gIniExe), "EXE\", "eKanban\Back\")
               .INI_String = gStr & eKanban_Prefix(eKanban_Proc) & "\"
               .PutIni
            End If
            iH1_BackDir(eKanban_Proc) = Trim(.INI_String)
            If Trim(iH1_BackDir(eKanban_Proc)) <> "" Then
               If Right(iH1_BackDir(eKanban_Proc), 1) <> "\" Then
                  iH1_BackDir(eKanban_Proc) = iH1_BackDir(eKanban_Proc) & "\"
               End If
            End If
            '------------------- ｼｽﾃﾑ共通設定
            .INI_Section = "JText"
            .INI_FileID = gIniExe & gIniName
            
            '-------- 受注ﾃｷｽﾄﾌｧｲﾙ名
            .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_" & eKanban_Orderer(eKanban_Proc) & "_ﾌｧｲﾙ名"
            .INI_Default = ""
            .GetIni
            If Trim(.INI_String) = "" Then
               If eKanban_Proc = 0 Then
                  .INI_String = "Juchu*"
               Else
                  .INI_String = "Juchu*"
               End If
               .PutIni
            End If
            iH1_TopFName(eKanban_Proc) = Trim(.INI_String)
            '-------- 保持期間
            .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_" & eKanban_Orderer(eKanban_Proc) & "_保持期間"
            .INI_Default = ""
            .GetIni
            If Trim(.INI_String) = "" Then
               .INI_String = "30"
               .PutIni
            End If
            gCnt1 = CisFun.Val2(.INI_String)
            If gCnt1 = 0 Then
               gCnt1 = 10
            End If
            iH1_Kikan(eKanban_Proc) = gCnt1
            iH1_Kikan(eKanban_Proc) = CisFun.RSetFld(iH1_Kikan(eKanban_Proc), iH1_Kikan(eKanban_Proc).MaxLength, "###")
            '-------- 変換リスト
            .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_" & eKanban_Orderer(eKanban_Proc) & "_変換リスト"
            .INI_Default = ""
            .GetIni
            If Trim(.INI_String) = "" Then
               .INI_String = "No"
               .PutIni
            End If
            H1ch_Prt(eKanban_Proc).Value = 2
            If UCase(Trim(.INI_String)) = "Y" Or UCase(Trim(.INI_String)) = "YES" Then
               H1ch_Prt(eKanban_Proc).Value = 1
            Else
               H1ch_Prt(eKanban_Proc).Value = 0
            End If
            '-------- 変換リスト(エラー内容)
            .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_" & eKanban_Orderer(eKanban_Proc) & "_変換リスト_エラー出力"
            .INI_Default = ""
            .GetIni
            If Trim(.INI_String) = "" Then
               .INI_String = "Yes"
               .PutIni
            End If
            If UCase(Trim(.INI_String)) = "Y" Or UCase(Trim(.INI_String)) = "YES" Then
               H1Ch_Err(eKanban_Proc).Value = 1
            Else
               H1Ch_Err(eKanban_Proc).Value = 0
            End If
            '-------- 変換リスト(重複)
            .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_" & eKanban_Orderer(eKanban_Proc) & "_変換リスト_重複"
            .INI_Default = ""
            .GetIni
            If Trim(.INI_String) = "" Then
               .INI_String = "Yes"
               .PutIni
            End If
            If UCase(Trim(.INI_String)) = "Y" Or UCase(Trim(.INI_String)) = "YES" Then
               H1Ch_Dbl(eKanban_Proc).Value = 1
            Else
               H1Ch_Dbl(eKanban_Proc).Value = 0
            End If
            '-------- 変換リスト(仮登録内容)
            .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_" & eKanban_Orderer(eKanban_Proc) & "_変換リスト_仮登録出力"
            .INI_Default = ""
            .GetIni
            If Trim(.INI_String) = "" Then
               .INI_String = "Yes"
               .PutIni
            End If
            If UCase(Trim(.INI_String)) = "Y" Or UCase(Trim(.INI_String)) = "YES" Then
               H1Ch_Kari(eKanban_Proc) = 1
            Else
               H1Ch_Kari(eKanban_Proc) = 0
            End If
            '-------- 変換リスト(その他)
            .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_" & eKanban_Orderer(eKanban_Proc) & "_変換リスト_その他出力"
            .INI_Default = ""
            .GetIni
            If Trim(.INI_String) = "" Then
               .INI_String = "Yes"
               .PutIni
            End If
            If UCase(Trim(.INI_String)) = "Y" Or UCase(Trim(.INI_String)) = "YES" Then
               H1Ch_Etc(eKanban_Proc) = 1
            Else
               H1Ch_Etc(eKanban_Proc) = 0
            End If
        End With
    Next eKanban_Proc

End Sub
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()

    DBPut = False
'  ( 確認メッセージ )
    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = "     【 テキスト連携 設定変更確認 】      "
        .MB_MSG(3) = "    テキスト連携 設定を変更しますか？     "
        .MB_Title = "テキスト連携設定"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With

    
'    For eKanban_Proc = 0 To 2
    For eKanban_Proc = 0 To 3                           '2010/04/19
        With CisFun
            '------------------- ｸﾗｲｱﾝﾄ個別設定
            .INI_Section = "JText"
            .INI_FileID = gIniExe & gwsIniName
            '-------- 受注ﾃｷｽﾄ格納先
            .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_" & eKanban_Orderer(eKanban_Proc) & "_ﾌｧｲﾙ格納先"
            .INI_Default = ""
            If Trim(iH1_Dir(eKanban_Proc)) <> "" Then
               If Right(iH1_Dir(eKanban_Proc), 1) <> "\" Then
                  iH1_Dir(eKanban_Proc) = iH1_Dir(eKanban_Proc) & "\"
               End If
            End If
            .INI_String = iH1_Dir(eKanban_Proc)
            .PutIni
            
            '-------- 受注ﾃｷｽﾄ退避先(SIMPR-A)
            .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_" & eKanban_Orderer(eKanban_Proc) & "_ﾌｧｲﾙ退避先"
            .INI_Default = ""
            If Trim(iH1_BackDir(eKanban_Proc)) <> "" Then
               If Right(iH1_BackDir(eKanban_Proc), 1) <> "\" Then
                  iH1_BackDir(eKanban_Proc) = iH1_BackDir(eKanban_Proc) & "\"
               End If
            End If
            .INI_String = iH1_BackDir(eKanban_Proc)
            .PutIni
            
            '------------------- ｼｽﾃﾑ共通設定
            .INI_Section = "JText"              ' 2009.09.15 upd : eKanban-->JText
            .INI_FileID = gIniExe & gIniName
            
            '-------- 受注ﾃｷｽﾄﾌｧｲﾙ名
            .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_" & eKanban_Orderer(eKanban_Proc) & "_ﾌｧｲﾙ名"
            .INI_Default = ""
            .INI_String = iH1_TopFName(eKanban_Proc)
            .PutIni
            '-------- 保持期間
            .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_" & eKanban_Orderer(eKanban_Proc) & "_保持期間"
            .INI_Default = ""
            .INI_String = Format(iH1_Kikan(eKanban_Proc), "0")
            .PutIni
            '-------- 変換リスト
            .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_" & eKanban_Orderer(eKanban_Proc) & "_変換リスト"
            .INI_Default = ""
            If H1ch_Prt(eKanban_Proc).Value = 1 Then
               .INI_String = "Yes"
            Else
               .INI_String = "No"
            End If
            .PutIni
            '-------- 変換リスト(エラー内容)
            .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_" & eKanban_Orderer(eKanban_Proc) & "_変換リスト_エラー出力"
            .INI_Default = ""
            If H1Ch_Err(eKanban_Proc).Value = 1 Then
               .INI_String = "Yes"
            Else
               .INI_String = "No"
            End If
            .PutIni
            '-------- 変換リスト(重複)
            .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_" & eKanban_Orderer(eKanban_Proc) & "_変換リスト_重複"
            .INI_Default = ""
            If H1Ch_Dbl(eKanban_Proc).Value = 1 Then
               .INI_String = "Yes"
            Else
               .INI_String = "No"
            End If
            .PutIni
            '-------- 変換リスト(仮登録内容)
            .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_" & eKanban_Orderer(eKanban_Proc) & "_変換リスト_仮登録出力"
            .INI_Default = ""
            If H1Ch_Kari(eKanban_Proc) = 1 Then
               .INI_String = "Yes"
            Else
               .INI_String = "No"
            End If
            .PutIni
            '-------- 変換リスト(その他)
            .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_" & eKanban_Orderer(eKanban_Proc) & "_変換リスト_その他出力"
            .INI_Default = ""
            If H1Ch_Etc(eKanban_Proc) = 1 Then
               .INI_String = "Yes"
            Else
               .INI_String = "No"
            End If
            .PutIni
        End With
    Next eKanban_Proc
        
' 更新確認
    MsgBox " " & vbCrLf & "         テキスト連携 設定更新終了 ！！        " & vbCrLf & " ", _
            vbOKOnly + vbInformation, "テキスト連携 設定変更"               '
'    B1_Area1.Enabled = True
    
    On Error GoTo 0
    
'    Call CisDB.DBDISConnect
    
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function

