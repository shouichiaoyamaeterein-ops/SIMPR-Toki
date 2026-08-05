VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Begin VB.Form CNJ0210 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "生産計画処理"
   ClientHeight    =   11115
   ClientLeft      =   405
   ClientTop       =   1245
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
   ForeColor       =   &H00004000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11115
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin VB.PictureBox ProcTime_PNL 
      Height          =   540
      Left            =   11670
      ScaleHeight     =   480
      ScaleWidth      =   3645
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   390
      Width           =   3705
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   847
         ForeColor       =   128
         Caption         =   "処理終了予定時刻"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D H1lb_EndTime 
         Height          =   480
         Left            =   2160
         Top             =   0
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   847
         ForeColor       =   128
         Caption         =   "ｈｈ時nn分"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
   End
   Begin VB.Timer Timer2 
      Left            =   495
      Top             =   1185
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00808080&
      Height          =   3765
      Left            =   4485
      ScaleHeight     =   3705
      ScaleWidth      =   6210
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3555
      Width           =   6270
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   3465
         Left            =   90
         Top             =   105
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   6112
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
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFC0&
            Height          =   1200
            Index           =   1
            Left            =   510
            ScaleHeight     =   1140
            ScaleWidth      =   5340
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1590
            Width           =   5400
            Begin VB.Label H1lb_ZTime 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "MM分 SS秒"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   1755
               TabIndex        =   19
               Top             =   855
               Width           =   1230
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "(処理時間)"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   225
               Index           =   3
               Left            =   405
               TabIndex        =   18
               Top             =   855
               Width           =   1140
            End
            Begin VB.Label H1lb_ZDate 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "YYYY年 MM月 DD日 HH:NN"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1755
               TabIndex        =   17
               Top             =   585
               Width           =   3075
            End
            Begin VB.Label H1lb_ZOutKB 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "XXXXXXXXXXXXXXXXXX"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1755
               TabIndex        =   16
               Top             =   330
               Width           =   2970
            End
            Begin VB.Label H1lb_ZYM 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "YYYY 年 MM 月度"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1755
               TabIndex        =   15
               Top             =   90
               Width           =   2025
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "(対象年月)"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   225
               Index           =   1
               Left            =   405
               TabIndex        =   14
               Top             =   105
               Width           =   1140
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "( 処理日 )"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   225
               Index           =   0
               Left            =   435
               TabIndex        =   13
               Top             =   600
               Width           =   1080
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "(処理内容)"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   225
               Index           =   2
               Left            =   405
               TabIndex        =   12
               Top             =   345
               Width           =   1140
            End
         End
         Begin VB.PictureBox MSG_Area 
            BackColor       =   &H00FFC0FF&
            Height          =   555
            Left            =   75
            ScaleHeight     =   495
            ScaleWidth      =   5775
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   2805
            Width           =   5835
            Begin Cis3D_v60.CIS3D lb_msg 
               Height          =   435
               Left            =   90
               Top             =   30
               Width           =   5595
               _ExtentX        =   9869
               _ExtentY        =   767
               BackColor       =   16761087
               Caption         =   "生産計画処理を行います"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   16761087
               cBoderColor2    =   16761087
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cBoderStyle     =   2
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D17 
            Height          =   1470
            Index           =   1
            Left            =   75
            Top             =   90
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   2593
            ForeColor       =   16711680
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
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cPositionX      =   150
            Begin CisYM_V60.CisYM iH1_YM 
               Height          =   360
               Left            =   2715
               TabIndex        =   1
               Top             =   480
               Width           =   1560
               _ExtentX        =   2646
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
               cContent        =   1
            End
            Begin VB.CheckBox H1chk_CLR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "前回内容削除"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐ明朝"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   270
               Left            =   4095
               TabIndex        =   2
               Top             =   1125
               Width           =   1680
            End
            Begin Cis3D_v60.CIS3D CIS3D1 
               Height          =   300
               Index           =   1
               Left            =   1005
               Top             =   510
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "今回対象年月"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   1
               cBoderStyle     =   2
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D10 
            Height          =   1200
            Left            =   75
            Top             =   1590
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   2117
            BackColor       =   0
            ForeColor       =   65535
            Caption         =   "前回内容"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   6
            cPositionX      =   15
            cPositionY      =   120
         End
      End
   End
   Begin VB.PictureBox C1_Area1 
      BackColor       =   &H80000007&
      Height          =   1485
      Left            =   10890
      ScaleHeight     =   1425
      ScaleWidth      =   3195
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5835
      Visible         =   0   'False
      Width           =   3255
      Begin Cis3D_v60.CIS3D UC_3D37 
         Height          =   435
         Index           =   0
         Left            =   30
         Top             =   510
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   767
         BackColor       =   12640511
         ForeColor       =   0
         Caption         =   "内示入力件数"
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
         cAlingnment     =   1
         cPositionX      =   360
         Begin Cis3D_v60.CIS3D C1lb_Input 
            Height          =   345
            Index           =   1
            Left            =   2190
            Top             =   30
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   609
            BackColor       =   14737632
            ForeColor       =   16576
            Caption         =   "999,999"
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
      Begin Cis3D_v60.CIS3D UC_3D41 
         Height          =   435
         Index           =   0
         Left            =   30
         Top             =   960
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   767
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "生産計画出力件数"
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
         cAlingnment     =   1
         cPositionX      =   120
         Begin Cis3D_v60.CIS3D C1lb_Input 
            Height          =   315
            Index           =   2
            Left            =   2190
            Top             =   60
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "999,999"
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
      Begin Cis3D_v60.CIS3D PFXLabel1 
         Height          =   465
         Index           =   0
         Left            =   30
         Top             =   30
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   820
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "【 処理対象件数 】"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
   Begin VB.Timer Timer1 
      Left            =   45
      Top             =   1200
   End
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'ｵﾌ固定
      Left            =   30
      TabIndex        =   0
      Text            =   "Dummy"
      Top             =   855
      Width           =   180
   End
   Begin VB.PictureBox Picture7 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   0
   End
   Begin Cis3D_v60.CIS3D PB_ENT 
      Height          =   645
      Left            =   8295
      Top             =   8280
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1138
      ForeColor       =   16576
      Caption         =   "処理開始"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
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
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   645
      Left            =   5475
      Top             =   8280
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1138
      ForeColor       =   16711680
      Caption         =   "終  了"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
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
   Begin Cis3D_v60.CIS3D PB_OK 
      Height          =   645
      Left            =   6885
      Top             =   8280
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1138
      ForeColor       =   65535
      Caption         =   "確  認"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
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
   Begin Cis3D_v60.CIS3D UC_3D11 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 生 産 計 画 処 理 】"
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
         Left            =   14355
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
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
         Height          =   285
         Left            =   11775
         Top             =   30
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
   Begin VB.Label BG_Logo 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  '透明
      Caption         =   "SIMPR-A"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   390
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5475
      TabIndex        =   8
      Top             =   8970
      Width           =   1485
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8295
      TabIndex        =   7
      Top             =   8970
      Width           =   1485
   End
   Begin VB.Label PNL_OK 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6885
      TabIndex        =   6
      Top             =   8970
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "CNJ0210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   生産計画処理
'**       フォームID    :   CNJ0210
'**       処理概要      :
'**
'**       作  成  日    :   2004/01/19  By CIS
'**       変  更  日    :   2005/02/10  Logo表示の修正（システム環境設定が反映されない！）
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim STime                   As Date
    Dim ETime                   As Date
    Dim ProcTime                As Long
    
    Const mGotColor             As Long = &HC0FFC0
    Const mLostColor            As Long = &HC0C0C0
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   # 二 重 起 動 防 止  #
'   #-------------------#
    If Not Dupli_Start Then End
'   #-------------------#
'   # 初 期 内 容 取 得  #
'   #-------------------#
    If Not IniGet Then End
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    Call CisFrm.InitFld
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
    
'   < シングルプロセスチェック >
    If Not SingleProcess_Check Then PB_END_Click
    
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   #     ロゴ表示       #
'   #-------------------#
'*---------------------------------------------------* ▽▽▽ 2005/02/10 Update Start ▽▽▽
'   CisFrm.BGLDisp BG_Logo, 10
    If gLogoDisp Then
        BG_Logo(0).Caption = gLogo
        CisFrm.BGLDisp BG_Logo, 10
    End If
'*---------------------------------------------------* △△△ 2005/02/10 Update End   △△△
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
'   < 内示対象年月獲得 >
    If Not Target_NJYM Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(1) = "    内示対象年月が確定していません。         "
            .MB_MSG(3) = "        区分名称 = '内示年月'   "
            .MB_MSG(4) = "                値 = '1'         "
            .MB_Button = Error
            .MBOX
            
            Call PB_END_Click
        End With
    End If
    
'   < 前回内容表示 >
    Call ProcRegistry_Check
    
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ｐａｉｎｔ イベント》                   +
'+----------------------------------------------------------+
Private Sub Form_Paint()
'   ( グラデーション )
    CisFrm.Gradation 0, 0, 255
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
    Key_Acc = False
    
    Select Case Key_Code
           Case vbKeyF12
                                If PB_END.Visible Then
                                    Call PB_END_Click                           '【終了】
                                Else
                                    If PB_OK.Visible Then Call PB_OK_Click      '【確認】
                                End If
           Case vbKeyReturn
                                If PB_ENT.Visible Then                          '【入力】
                                    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                    Call PB_ENT_Click
                                    Dummy.Enabled = False
                                End If
           Case Else:           Exit Function
    End Select
    
    Key_Acc = True
End Function
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    Call SingleProcess_Release
    Call CisDB.DBDISConnect                         ' ﾃﾞｰﾀﾍﾞｰｽ切断
    Call CisFrm.UnLoadDisp                          ' ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
    End
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Call ReturnPress
End Sub
'*----------------------------------*
'*      確　認　終　了　キ　ー      *
'*----------------------------------*
Private Sub PB_OK_Click()
    Call PB_END_Click
End Sub
'*----------------------------------*
'*      前 回 内 容 削 除 区 分
'*----------------------------------*
Private Sub H1chk_CLR_Click()
    If H1chk_CLR.Value = 1 Then
       H1chk_CLR.BackColor = mGotColor
    Else
       H1chk_CLR.BackColor = mLostColor
    End If
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*******************************
'*       メッセージ表示         *
'*******************************
Private Sub Timer2_Timer()
    CisFrm.MSGRelate
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "PR"
        Call DispChange(ProcHB)
        DoEvents
        
'       ( Proc Start )
        Call Proc_Start
    
        ProcHB = "C1"
        Call DispChange(ProcHB)
        
        With CisFrm
            Set .MSGObject = lb_msg
            Set .MSGTimer = Timer2
            .MSGInterval = 100
            .MSGStopCount = 20
            .MSGType = Left_To_Right
            .MSGText = "【 処理終了 】確認して下さい"
            
            .MSGRelate
        End With
        
        GoTo ReturnPress_Ed
    End If
    
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
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
    Dim H1Mode          As Boolean
    Dim H1Color         As Long
    Dim PRMode          As Boolean
    Dim PRColor         As Long
    Dim C1Mode          As Boolean
    Dim C1Color         As Long
    
    H1Mode = False:     H1Color = gPLostSel
    PRMode = False:     PRColor = gPLostSel
    C1Mode = False:     C1Color = gPLostSel
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "PR":      PRMode = True:  PRColor = gPGotSel
        Case "C1":      C1Mode = True:  C1Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    C1_Area1.Visible = C1Mode
    
    PB_ENT.Visible = Not (PRMode Or C1Mode)
    PNL_ENT.Visible = Not (PRMode Or C1Mode)
    PB_END.Visible = Not (PRMode Or C1Mode)
    PNL_END.Visible = Not (PRMode Or C1Mode)
    ProcTime_PNL.Visible = PRMode
    
    PB_OK.Visible = C1Mode
    PNL_OK.Visible = C1Mode
    
    Dummy.Enabled = False
End Function
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
'   < 対象年月 >
    If RTrim$(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    今回対象年月を入力して下さい         "
            .MB_Button = Error
            .MBOX
            iH1_YM.SetFocus
            Exit Function
        End With
    End If
    
'   < 前回内容 >
    Call ItemsClearSKZ
    gSL_Select = "Select Distinct 内示年月 From 生産計画前回テーブル"
    If SKZRead(gSL_Select, 1) Then
        If SKZ.内示年月 <> iH1_YM Then H1chk_CLR.Value = 1
    End If
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "    生産計画処理を開始します？         "
        .MB_Button = Yes_No
        If Not .MBOX Then Exit Function
    End With
    
    Head1Chk = True
End Function
'*--------------------------*
'*      処　理　開　始      *
'*--------------------------*
Private Function Proc_Start() As Boolean
    Proc_Start = False
    
    If ProcTime <> 0 Then
        STime = DateAdd("S", ProcTime, Now())
        H1lb_EndTime = Format(STime, "hh時 nn分")
    End If
    
    STime = Now()
    
    lb_msg = "【 処理中 】しばらくお待ち下さい": DoEvents
    
    On Error GoTo Proc_Start_Err
    Call CisDB.DBTran(TransBegin)

    With CisDB
        .SQL = "生産計画処理"
        .StoadoCount = 4
        .ParaNo = 0:    .ParaIO = Return用
        .ParaNo = 1:    .ParaIO = Input用
        .ParaNo = 2:    .ParaIO = Input用
        .ParaNo = 3:    .ParaIO = OutPut用
        .ParaNo = 4:    .ParaIO = OutPut用
    
        .ParaNo = 1:    .ParaValue = iH1_YM
        .ParaNo = 2:    .ParaValue = Abs(CisFun.Val2(H1chk_CLR.Value) - 1)
    
        .DBStored
    
        .ParaNo = 0
        If .ParaValue = 1 Or .ParaValue = 3 Or .ParaValue = 5 Then
            CisFun.MB_Lines = 5
            CisFun.MB_MSG(2) = "    前倒し日数が３ヶ月を超えるデータがあります。        "
            CisFun.MB_Button = OK
            CisFun.MBOX
        End If
    
        .ParaNo = 3:    C1lb_Input(1) = Format(.ParaValue, "#,###")
        .ParaNo = 4:    C1lb_Input(2) = Format(.ParaValue, "#,###")
    End With
    
    If CisFun.Val2(H1chk_CLR.Value) = 0 Then
'       < 内示テーブル：処理区分クリア >
        gSL_Select = "UpDate 内示テーブル Set"
        gSL_Select = gSL_Select & " 処理区分 = 0,"
        gSL_Select = gSL_Select & " 更新者 = '" & gTantoName & "',"
        gSL_Select = gSL_Select & " 更新日 = GetDate(),"
        gSL_Select = gSL_Select & " 更新端末 = Substring( Host_Name(), 1, 20 )"
        gSL_Select = gSL_Select & " Where 内示年月 = '" & iH1_YM & "'"
        gSL_Select = gSL_Select & " And   処理区分 <> 0"
        
        With CisDB
            .SQL = gSL_Select
            .DBExec
        End With
    End If

    ETime = Now()

'   < 今回内容保存 >
    Call ProcRegistry_Save
    
    Call CisDB.DBTran(TransCommit)
    On Error GoTo 0
    
    Proc_Start = True
    Exit Function

Proc_Start_Err:
    CisFun.ErrorBox "生産計画処理"
    End
End Function
'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'▽
'▽                         名　称　マ　ス　タ　確　認  ( シ ス テ ム 共 通 定 数 等 )
'▽
'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'*------------------------------------------------*
'*          内　示　対　象　年　月　獲　得          *
'*------------------------------------------------*
Private Function Target_NJYM()
    Target_NJYM = False
    
'   < 内示対象年月 >
    Call GetNJProc_YM(1)

'   < 生産計画年月 >
    Call GetNJProc_YM(2)

    If RTrim$(gn_NJYM) > RTrim$(gn_SKYM) Then
        gn_SKYM = gn_NJYM
        H1chk_CLR.Value = 1
        Call H1chk_CLR_Click
    End If
    
'( 2004.06.29 MNT START )
'    If RTrim$(gn_SKYM) = "" Then Exit Function
    If RTrim$(gn_SKYM) = "" Then gn_SKYM = gn_NJYM
'( 2004.06.29 MNT  END  )
    
    iH1_YM = gn_SKYM
    Target_NJYM = True
End Function
'*---------------------------------------------------------*
'*                  前　回　内　容　獲　得                  *
'*---------------------------------------------------------*
Private Sub ProcRegistry_Check()
    
    Call ItemsClearSYMCreate
    
    With SYMCreate
        .区分名称 = "処理内容"
        .区分桁数 = 10
        .値 = App.EXEName
        .制御区分 = "X"
        .値名称 = "生産計画処理"
        .作成者 = gTanto
    End With
    
    Call SYMInitCreate(False)

    Call ItemsClearSYM
    gSL_Select = "Select * From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = '処理内容'"
    gSL_Select = gSL_Select & " And   区分タイプ = ''"
    gSL_Select = gSL_Select & " And   値 = '" & App.EXEName & "'"
    Call SYMRead(gSL_Select, 1)
    
    With SYM
        If RTrim$(.英数字1) <> "" Then
            H1lb_ZYM = Format(.英数字1, "@@@@年 @@月度")
            If .数字1 = 1 Then H1lb_ZOutKB = "前回内容削除"
        End If
        If RTrim$(.英数字2) <> "" Then
            H1lb_ZDate = ""
            H1lb_ZDate = H1lb_ZDate & Format(.英数字2, "@@@@年 @@月 @@日 ")
            H1lb_ZDate = H1lb_ZDate & Format(.英数字3, "@@:@@")
        End If
        
        ProcTime = .数字2
        If .数字2 <> 0 Then
            STime = 0
            STime = DateAdd("S", .数字2, STime)
            
            If CisFun.Val2(DatePart("H", STime)) <> 0 Then
                gStr = "HH時間 NN分 SS秒"
            Else
                If CisFun.Val2(DatePart("N", STime)) <> 0 Then
                    gStr = "NN分 SS秒"
                Else
                    gStr = "SS秒"
                End If
            End If
            H1lb_ZTime = Format(STime, gStr)
        End If
    End With

End Sub
'*---------------------------------------------------------*
'*                  今　回　内　容　保  存                  *
'*---------------------------------------------------------*
Private Sub ProcRegistry_Save()

    Call ItemsClearSYM
    
    With SYM
        .区分名称 = "処理内容"
        .区分桁数 = 10
        .値 = App.EXEName
        .制御区分 = "X"
        .値名称 = "生産計画処理"
        .英数字1 = iH1_YM
        .数字1 = CisFun.Val2(H1chk_CLR.Value)
        .数字2 = DateDiff("S", STime, ETime)
        .英数字2 = Format(Now(), "YYYYMMDD")
        .英数字3 = Format(Now(), "HHNN")
        .作成者 = gTanto
    End With

    Call SYMUpdate

    Call ItemsClearSYM
    
    With SYM
        .区分名称 = "内示年月"
        .区分桁数 = 1
        .値 = 2
        .制御区分 = "X"
        .値名称 = "生産計画年月"
        .英数字1 = iH1_YM
        .作成者 = gTanto
    End With
    
    Call SYMUpdate
    
End Sub
'*---------------------------------------------------------*
'*          シ ン グ ル プ ロ セ ス チ ェ ッ ク             *
'*---------------------------------------------------------*
Private Function SingleProcess_Check()
    SingleProcess_Check = False
    
    Call ItemsClearSYMCreate
    
    With SYMCreate
        .区分名称 = "シングルプロセス"
        .区分桁数 = 10
        .値 = App.EXEName
        .制御区分 = "X"
        .値名称 = "生産計画処理"
        .作成者 = gTanto
    End With
    
    Call SYMInitCreate(False)
    
    Call ItemsClearSYM
    gSL_Select = "Select 英数字1 From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称   = 'シングルプロセス'"
    gSL_Select = gSL_Select & " And   区分タイプ = ''"
    gSL_Select = gSL_Select & " And   値        = '" & App.EXEName & "'"
    Call SYMRead(gSL_Select, 1)
    
    With CisFun
        If RTrim$(SYM.英数字1) <> "" And RTrim$(SYM.英数字1) <> .WSID Then
            .MB_Lines = 3
            .MB_MSG(1) = "    他のパソコンで起動中です         "
            .MB_MSG(2) = "        ( " & RTrim$(SYM.英数字1) & " )"
            .MB_Button = Error
            .MBOX
            Exit Function
        End If
    End With

'   < プロセスのロック >
    gSL_Select = "UpDate 名称マスタ Set"
    gSL_Select = gSL_Select & " 英数字1 = SubString( Host_Name(), 1, 20 ),"
    gSL_Select = gSL_Select & " 更新日 = GETDATE(),"
    gSL_Select = gSL_Select & " 更新者 = '" & RTrim(gTanto) & "',"
    gSL_Select = gSL_Select & " 更新端末 = SubString( Host_Name(), 1, 20 )"
    gSL_Select = gSL_Select & " Where 区分名称 = 'シングルプロセス'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   And 値 = '" & App.EXEName & "'"

    With CisDB
        .SQL = gSL_Select
        .DBExec
    End With

    SingleProcess_Check = True
End Function
'*-------------------------------------------------*
'*          シ ン グ ル プ ロ セ ス 解 除           *
'*-------------------------------------------------*
Private Function SingleProcess_Release()
    SingleProcess_Release = False
    
    gSL_Select = "UpDate 名称マスタ Set"
    gSL_Select = gSL_Select & " 英数字1 = '',"
    gSL_Select = gSL_Select & " 更新日 = GETDATE(),"
    gSL_Select = gSL_Select & " 更新者 = '" & RTrim(gTanto) & "',"
    gSL_Select = gSL_Select & " 更新端末 = SubString( Host_Name(), 1, 20 )"
    gSL_Select = gSL_Select & " Where 区分名称 = 'シングルプロセス'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   And 値 = '" & App.EXEName & "'"

    With CisDB
        .SQL = gSL_Select
        .DBExec
    End With
    
    SingleProcess_Release = True
End Function

