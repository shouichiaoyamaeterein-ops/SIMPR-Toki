VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{490AF35D-478A-11D4-AEC2-00000EA57F9E}#4.0#0"; "CISMSG~2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Begin VB.Form CXJ0030 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "受注かんばん入力"
   ClientHeight    =   10005
   ClientLeft      =   345
   ClientTop       =   855
   ClientWidth     =   13980
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
   ScaleHeight     =   10005
   ScaleWidth      =   13980
   WindowState     =   2  '最大化
   Begin VB.PictureBox H1_Area1 
      Height          =   1350
      Left            =   1170
      ScaleHeight     =   1290
      ScaleWidth      =   6825
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1755
      Width           =   6885
      Begin Cis3D_v60.CIS3D CIS3D17 
         Height          =   1215
         Left            =   60
         Top             =   45
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   2143
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
         cAlingnment     =   6
         cPositionY      =   60
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
            Left            =   975
            TabIndex        =   1
            Top             =   450
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
            Left            =   975
            TabIndex        =   5
            Top             =   810
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
            Left            =   45
            TabIndex        =   4
            Top             =   810
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
            Left            =   45
            TabIndex        =   0
            Top             =   450
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
            Left            =   1905
            TabIndex        =   2
            Top             =   450
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
            Left            =   2835
            TabIndex        =   7
            Top             =   810
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
            Left            =   2835
            TabIndex        =   3
            Top             =   450
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
            Left            =   1905
            TabIndex        =   6
            Top             =   810
            Width           =   840
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1215
         Index           =   0
         Left            =   4260
         Top             =   45
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   2143
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
            TabIndex        =   8
            Top             =   570
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
            Left            =   435
            Top             =   60
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
            TabIndex        =   9
            Top             =   570
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
   Begin Cis3D_v60.CIS3D PNL_PF 
      Height          =   855
      Left            =   8370
      Top             =   1740
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1508
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderStyle     =   1
      Begin Cis3D_v60.CIS3D PB_END 
         Height          =   765
         Left            =   3660
         Top             =   45
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   1349
         ForeColor       =   12582912
         Caption         =   " 終了 (F12)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   6
         cAlingnment     =   6
         cButton         =   -1  'True
      End
      Begin Cis3D_v60.CIS3D PB_ENT 
         Height          =   765
         Left            =   1860
         Top             =   45
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   1349
         ForeColor       =   16576
         Caption         =   "開始(Enter)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   6
         cAlingnment     =   6
         cButton         =   -1  'True
      End
      Begin Cis3D_v60.CIS3D PB_ENT2 
         Height          =   765
         Left            =   60
         Top             =   45
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   1349
         ForeColor       =   0
         Caption         =   "枚数初期化(F8)"
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
         cBoderWidth     =   6
         cAlingnment     =   6
         cButton         =   -1  'True
         cPositionY      =   40
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   945
      Top             =   1455
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox PNL_Start 
      BackColor       =   &H00000000&
      Height          =   825
      Left            =   3825
      ScaleHeight     =   765
      ScaleWidth      =   7125
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   795
      Visible         =   0   'False
      Width           =   7185
      Begin CisMsgRelate_v60.CISMsgRelate ProcMsg 
         Height          =   645
         Left            =   75
         Top             =   60
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   1138
         BackColor       =   65535
         Caption         =   "かんばんソータの【開始】キーを押して下さい"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   15.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         mrInterval      =   50
         mrStopCnt       =   40
      End
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   6465
      Left            =   1185
      ScaleHeight     =   6405
      ScaleWidth      =   12660
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3825
      Width           =   12720
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   6285
         Left            =   60
         Top             =   60
         Width           =   12540
         _ExtentX        =   22119
         _ExtentY        =   11086
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
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   6195
            Index           =   0
            Left            =   45
            Top             =   45
            Width           =   12450
            _ExtentX        =   21960
            _ExtentY        =   10927
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
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   480
               Index           =   0
               Left            =   120
               Top             =   3105
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   847
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "№"
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   480
               Index           =   12
               Left            =   450
               Top             =   3105
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   847
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "7"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
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
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   480
               Index           =   13
               Left            =   2415
               Top             =   3105
               Width           =   1980
               _ExtentX        =   3493
               _ExtentY        =   847
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "8"
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
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   480
               Index           =   14
               Left            =   4395
               Top             =   3105
               Width           =   1980
               _ExtentX        =   3493
               _ExtentY        =   847
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "9"
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
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   480
               Index           =   15
               Left            =   6375
               Top             =   3105
               Width           =   1980
               _ExtentX        =   3493
               _ExtentY        =   847
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "10"
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
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   480
               Index           =   16
               Left            =   8355
               Top             =   3105
               Width           =   1980
               _ExtentX        =   3493
               _ExtentY        =   847
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "11"
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
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   480
               Index           =   17
               Left            =   10335
               Top             =   3105
               Width           =   1980
               _ExtentX        =   3493
               _ExtentY        =   847
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "12"
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
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   495
               Index           =   18
               Left            =   120
               Top             =   3585
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   873
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "1"
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   495
               Index           =   19
               Left            =   120
               Top             =   4080
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   873
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "2"
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   495
               Index           =   20
               Left            =   120
               Top             =   4575
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   873
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "3"
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   495
               Index           =   21
               Left            =   120
               Top             =   5070
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   873
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "4"
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   495
               Index           =   22
               Left            =   120
               Top             =   5565
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   873
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "5"
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   480
               Index           =   50
               Left            =   120
               Top             =   135
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   847
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "№"
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   480
               Index           =   3
               Left            =   2415
               Top             =   135
               Width           =   1980
               _ExtentX        =   3493
               _ExtentY        =   847
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "2"
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
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   480
               Index           =   4
               Left            =   4395
               Top             =   135
               Width           =   1980
               _ExtentX        =   3493
               _ExtentY        =   847
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "3"
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
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   480
               Index           =   5
               Left            =   6375
               Top             =   135
               Width           =   1980
               _ExtentX        =   3493
               _ExtentY        =   847
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "4"
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
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   480
               Index           =   6
               Left            =   8355
               Top             =   135
               Width           =   1980
               _ExtentX        =   3493
               _ExtentY        =   847
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "5"
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
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   480
               Index           =   7
               Left            =   10335
               Top             =   135
               Width           =   1980
               _ExtentX        =   3493
               _ExtentY        =   847
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "6"
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
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   495
               Index           =   2
               Left            =   120
               Top             =   615
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   873
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "1"
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D B1_ErrName 
               Height          =   495
               Index           =   1
               Left            =   465
               Top             =   615
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   873
               BackColor       =   16761087
               ForeColor       =   0
               Caption         =   "読取不能"
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   1
               cPositionX      =   30
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   495
               Index           =   8
               Left            =   120
               Top             =   1110
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   873
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "2"
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D B1_ErrName 
               Height          =   495
               Index           =   2
               Left            =   465
               Top             =   1110
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   873
               BackColor       =   16761087
               ForeColor       =   0
               Caption         =   "重複読取"
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   1
               cPositionX      =   30
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   495
               Index           =   9
               Left            =   120
               Top             =   1605
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   873
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "3"
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   495
               Index           =   10
               Left            =   120
               Top             =   2100
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   873
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "4"
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   495
               Index           =   11
               Left            =   120
               Top             =   2595
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   873
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "5"
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   480
               Index           =   1
               Left            =   450
               Top             =   135
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   847
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "1"
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
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin vsFlexLib.vsFlexArray vsGrid 
               Height          =   6000
               Left            =   90
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   90
               Width           =   12255
               _Version        =   196608
               _ExtentX        =   21616
               _ExtentY        =   10583
               _StockProps     =   228
               Appearance      =   1
               ConvInfo        =   1418783674
               FormatString    =   "№|XXXXXXXXX1X|XXX4|XXXXXXXXX1X|XXX4|XXXXXXXXX1X|XXX4|XXXXXXXXX1X|XXX4|XXXXXXXXX1X|XXX4|XXXXXXXXX1X|XXX4|"
               Rows            =   12
               Cols            =   20
               BackColor       =   12648384
               BackColorBkg    =   12632256
               FocusRect       =   0
               HighLight       =   0
               ScrollBars      =   0
               RowHeightMin    =   500
               GridLines       =   2
               SelectionMode   =   1
               Begin Cis3D_v60.CIS3D B1_ErrName 
                  Height          =   495
                  Index           =   3
                  Left            =   330
                  Top             =   1485
                  Width           =   1410
                  _ExtentX        =   2487
                  _ExtentY        =   873
                  BackColor       =   16761087
                  ForeColor       =   0
                  Caption         =   "ﾁｪｯｸｴﾗｰ"
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
                  cFont3DColor1   =   8421504
                  cFont3DColor2   =   -2147483643
                  cFont3D         =   2
                  cAlingnment     =   1
                  cPositionX      =   30
               End
               Begin Cis3D_v60.CIS3D B1_ErrName 
                  Height          =   495
                  Index           =   4
                  Left            =   330
                  Top             =   1980
                  Width           =   1410
                  _ExtentX        =   2487
                  _ExtentY        =   873
                  BackColor       =   16761087
                  ForeColor       =   0
                  Caption         =   "ﾎﾟｹｯﾄｵｰﾊﾞｰ"
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
                  cFont3DColor1   =   8421504
                  cFont3DColor2   =   -2147483643
                  cFont3D         =   2
                  cAlingnment     =   1
                  cPositionX      =   30
               End
            End
         End
      End
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
      Caption         =   "【 受注かんばん入力 】"
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
         Height          =   420
         Left            =   -30
         Top             =   -30
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   741
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "デモ"
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
         cPositionY      =   10
      End
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14355
         Top             =   75
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
         Left            =   11775
         Top             =   30
         Width           =   2550
         _ExtentX        =   4498
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
   Begin Cis3D_v60.CIS3D PNL_Kensu 
      Height          =   630
      Left            =   7410
      Top             =   3150
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1111
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
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   570
         Left            =   30
         Top             =   30
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   1005
         ForeColor       =   0
         Caption         =   "入力枚数"
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
         Begin Cis3D_v60.CIS3D lb_InCnt 
            Height          =   360
            Left            =   1065
            Top             =   105
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            BackColor       =   14737632
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
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
            cPositionX      =   -80
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   570
         Left            =   2175
         Top             =   30
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   1005
         BackColor       =   16744576
         ForeColor       =   0
         Caption         =   "正常枚数"
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
         Begin Cis3D_v60.CIS3D lb_OKCnt 
            Height          =   360
            Left            =   1065
            Top             =   105
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            BackColor       =   16761024
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
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
            cPositionX      =   -80
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D10 
         Height          =   570
         Left            =   4320
         Top             =   30
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   1005
         BackColor       =   16744703
         ForeColor       =   0
         Caption         =   "異常枚数"
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
         Begin Cis3D_v60.CIS3D lb_ErrCnt 
            Height          =   360
            Left            =   1065
            Top             =   105
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            BackColor       =   16761087
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
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
            cPositionX      =   -80
         End
      End
   End
   Begin Cis3D_v60.CIS3D lb_ErrMsg 
      Height          =   930
      Left            =   1275
      Top             =   555
      Visible         =   0   'False
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   1640
      BackColor       =   255
      Caption         =   "かんばん読取エラーメッセージ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
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
   Begin VB.Timer Timer2 
      Left            =   1920
      Top             =   885
   End
   Begin VB.Timer Timer1 
      Left            =   1650
      Top             =   885
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
      IMEMode         =   3  'ｵﾌ固定
      Left            =   1350
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   885
      Width           =   225
   End
End
Attribute VB_Name = "CXJ0030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   受注かんばん入力
'**       フォームID    :   CXJ0030
'**       処理概要      :
'**
'**       作  成  日    :   2004/03/30  By CIS
'**       変  更  日    :   2004/06/24  By CIS - 受領書・出荷時発行の対応
'**       変  更  日    :   2005/11/07  By CIS オプション追加（受注かんばん読取ｴﾗｰﾘｽﾄ発行有無）
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    
    Dim mRId                    As Long
    Dim mCId                    As Long
    
    Dim mShiwake                As Integer

    Dim mGotColor       As Long
    Dim mLostColor      As Long
    
    Dim mMColor         As Long
    Dim mNColor         As Long
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
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止
    
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
    
    mMColor = &HFFFFC0
    mNColor = &HFFC0C0
    
'    mGotColor = OpH1_Group(0).BackColor
'    mLostColor = OpH1_Group(1).BackColor
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
    H1Op_JKbn(1).Value = True
    
    mShiwake = 1    ' 品番仕分け
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
'   >> 名称マスタ確認
    Call MeisyoInit
'   >> 日付表示
    iH1_Ymd = Format(Date, "yyyymmdd")
    iH1_Bin = "01"
    
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
' かんばん処理 初期設定
    Set TX597G_Comm = MSComm1
    Call TX597GInfoGet
    kBas_Proc = "J"
    Call cKanbanInit
    
    If kBas_ExecMode Then
       K_Sykbnm.Visible = True
    Else
       K_Sykbnm.Visible = False
    End If
    
    '発行元
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 4
    Call IRNRead("", 1)
    Qr_HMoto = RTrim(Mid(IRN.oyian, 1, 7))
    If Trim(Qr_HMoto) = "" Then Qr_HMoto = "SIMPR-A"
    
    Call GridClear
    
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
    Key_Acc = False
    
    Select Case Key_Code
           Case vbKeyS:         If kBas_ExecMode Then kBas_DemoFlg = "S"
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                                If PB_Ent.Visible Then                      '【入力】
                                    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                    Call PB_ENT_Click
                                    Dummy.Enabled = False
                                End If
           Case vbKeyF8
                                If PB_ENT2.Visible Then                      '【入力】
                                    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                    Call PB_ENT2_Click
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
    
    Call End_Proc
    
    Call CisDB.DBDISConnect                            ' ﾃﾞｰﾀﾍﾞｰｽ切断
    Call CisFrm.UnLoadDisp                          ' ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
    End
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Call ReturnPress
End Sub
'+----------------------------+
'+    枚数初期化(F8)           +
'+----------------------------+
Private Sub PB_ENT2_Click()
    With CisFun
         .MB_Lines = 4
         .MB_MSG(2) = "     読取枚数の初期化を行います！      "
         .MB_MSG(4) = "     　　　よろしいですか？           "
         .MB_Title = "初期化"
         .MB_Button = OK_CAN
         If Not .MBOX Then Exit Sub
    End With
    Call GridClear
    Erase kBas_TKey
    Erase kBas_TKey2
    Erase kBas_TMai
    
    ReDim kBas_TKey(kBas_PoketMax)
    ReDim kBas_TKey2(kBas_PoketMax)
    ReDim kBas_TMai(kBas_PoketMax)
    
    kBas_ReadFlg = False
    PB_ENT2.Visible = False
    lb_InCnt = ""
    lb_OKCnt = ""
    lb_ErrCnt = ""
End Sub
'+----------------------------+
'+      タイトルバー編集       +
'+----------------------------+
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'+----------------------------+
'+  かんばん読取ｴﾗｰﾒｯｾｰｼﾞ表示  +
'+----------------------------+
Private Sub Timer2_Timer()
    If lb_ErrMsg.Visible Then
       lb_ErrMsg.Visible = False
       Timer2.Interval = 600
    Else
       lb_ErrMsg.Visible = True
       Timer2.Interval = 2000
    End If
End Sub
'+----------------------------------------+
'+  H1Op_Jkbn ( Click )      受注区分
'+----------------------------------------+
Private Sub H1Op_JKbn_Click(Index As Integer)
    H1Op_JKbn(1).Tag = Index
    For Each gObj In H1Op_JKbn
       gObj.BackColor = &H8000000F
    Next gObj
    H1Op_JKbn(Index).BackColor = &H8080FF
    Select Case Index
        Case 1:     Back_Ymd.Caption = "納 入 日"
        Case Else:  Back_Ymd.Caption = "納 入 日"
    End Select
End Sub
'+-----------------------------+
'+  iH1_Ymd - 受注日・納入日
'+-----------------------------+
Private Sub iH1_Ymd_GotFocus()
    Dummy.Tag = ActiveControl.Name
End Sub
Private Sub iH1_Ymd_LostFocus()
    If ActiveControl.Name <> "Dummy" Then
        Dummy.Tag = ""
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
    Dim RDMode      As Boolean
    Dim RDColor     As Long
    H1Mode = False
    H1Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
    RDMode = False
    RDColor = gPLostSel
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
       Case "RD"
          RDMode = True
          RDColor = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.BackColor = B1Color
    PNL_PF.Visible = Not RDMode
    PB_ENT2.Visible = kBas_ReadFlg
    
' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)

    If Not Head1Chk Then GoTo ReturnPress_Ed
    
    Call KanbanRead
    
    ProcHB = "H1"
    Call DispChange(ProcHB)
    
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

    If RTrim(iH1_Ymd) = "" Then iH1_Ymd = Format(Now(), "yyyymmdd")
    If CisFun.Val2(iH1_Bin) = 0 Then iH1_Bin = "01"
'   << 受注日 or 納入日 >>
    If RTrim(iH1_Ymd) = "" Then
        CisFun.MB_MSG(2) = "    " & Replace(Back_Ymd.Caption, " ", "") & "を入力して下さい。      "
        CisFun.MBOX
        iH1_Ymd.SetFocus
        Exit Function
    Else
        If Not iH1_Ymd.cChkResult Then
            CisFun.MB_MSG(2) = "    " & Replace(Back_Ymd.Caption, " ", "") & "を正しく入力して下さい            "
            CisFun.MBOX
            iH1_Ymd.SetFocus
            Exit Function
        End If
    End If
'   << 便 >>
    If Trim(iH1_Bin) = "" Then
        CisFun.MB_MSG(2) = "    便を入力して下さい            "
        CisFun.MBOX
        iH1_Bin.SetFocus
        Exit Function
    End If

    With CisFun
        gStr = ""
        Select Case Kdou_Check(iH1_Ymd)
            Case 1: gStr = "      指定日付は稼働日ではありません          "
            Case 2: gStr = "      カレンダマスタ未登録                "
        End Select
    
        If Trim(gStr) <> "" Then
            .MB_MSG(2) = gStr
            .MBOX
            iH1_Ymd.SetFocus
            Exit Function
        End If
    End With

    Head1Chk = True
End Function
'//////////////////////////////////////////
'//
'//     かんばんソーター読取
'//
'//////////////////////////////////////////
Private Sub KanbanRead()

    Call CisFrm.MousePT(11)
' エラーﾒｯｾｰｼﾞ 消去
    Timer2.Enabled = False
    lb_ErrMsg.Visible = False
' 開始確認
    ProcHB = "RD"
    Call DispChange(ProcHB)
    
    kBas_ReadMsg = "ReadStart"

    CKN0000.Show vbModal
    
    If Not kBas_ReadOk Then GoTo KanbanRead_Ed
' 通信ﾎﾟｰﾄｵｰﾌﾟﾝ
    If Not TX597GOpen Then
        Call PB_END_Click
    End If
' デモ用設定
    kBas_DemoFno = 0
'///  かんばん読取処理 開始  １枚目
    Call ReadStart
'///
KanbanRead_Ed:
' 通信ﾎﾟｰﾄｸﾛｰｽﾞ
    Call TX597GClose
    
    ProcMsg.MsgActED
    ProcHB = "H1"
    Call DispChange(ProcHB)
    Call CisFrm.MousePT(1)
End Sub
'****************************************
'*  かんばん読取処理 開始  １枚目         *
'****************************************
Private Sub ReadStart()
    ProcMsg = "かんばんソータの【開始】キーを押して下さい"
    ProcMsg.MsgActST
    PNL_Start.Visible = True
    kBas_DemoFlg = ""
    ' デモ用
    If kBas_ExecMode Then
       Do Until kBas_DemoFlg = "S"
          DoEvents
       Loop
    End If
    DoEvents
' かんばんｿｰﾀｰ処理
' 初期処理
    Call TX597GInit
' かんばん読取
    Call TX597GRead

    Do Until Not ReadNext           ' 2枚目移行
    Loop
End Sub
'****************************************
'*  かんばん読取処理 開始  ２枚目以降     *
'****************************************
Private Function ReadNext()
    Static wPNo     As Integer
    
    ReadNext = True
    
    If Not TX597GComm Then GoTo ReadNext_ED
        
' 終了ｽﾃｰﾀｽ確認
    If Mid(kBas_CommData, 1, 1) = Chr(&H1B) And _
       Mid(kBas_CommData, 2, 1) = "E" Then
       Call ReadEnd
       GoTo ReadNext_ED
    End If
' かんばん読取中ﾒｯｾｰｼﾞ表示
    ProcMsg.MsgActED
    ProcMsg = "【 かんばん読取中 】"
    DoEvents
' かんばんソータｽﾃｰﾀｽﾁｪｯｸ
    kBas_ErrSts = TX597GCheck
    
    Select Case kBas_ErrSts
       Case 0
            Call KanbanProc         ' かんばん情報 処理 (設定値 = 指定ポケット№)
       Case 1
            Call TX597GPoket         ' ｿｰﾀｰﾎﾟｹｯﾄ指定
            Call TX597GRead          ' 次のかんばん読取
       Case Else
            Call ErrDisp("E")
            ReadNext = False
            Exit Function           ' ソーターエラー処理
    End Select
    
    Exit Function

ReadNext_ED:
    Call ErrDisp("")
    ReadNext = False
End Function
'*******************************
'*  かんばん読取処理 終了        *
'*******************************
Private Sub ReadEnd()
' 読取中ﾒｯｾｰｼﾞ消去
    PNL_Start.Visible = False
    DoEvents
' 読取確認
    kBas_ReadMsg = "ReadEnd"

' 通信ポートクローズ
    Call TX597GClose
    PNL_Start.Visible = False

'    CKN0000.Show vbModal
'    If Not kBas_ReadOk Then
'' 通信ポートクローズ
'       Call CommOpCl(MSComm1, False)
'       PNL_Start.Visible = False
'       Exit Sub
'    End If
'
'    Call ReadStart
End Sub
'********************************
'*  かんばん読取エラー処理 終了   *
'********************************
Private Sub ErrDisp(ErType As String)
    
    If PNL_PF.Visible Then Exit Sub
    
' 通信ﾎﾟｰﾄｸﾛｰｽﾞ
    Call TX597GClose
   
    If ErType = "E" Then
        If InStr(1, kBas_ErrMsg, vbCr) <> 0 Then
            lb_ErrMsg.cPositionY = 60
            lb_ErrMsg.cAlingnment = 7
        Else
            lb_ErrMsg.cPositionY = 0
            lb_ErrMsg.cAlingnment = 6
        End If
        lb_ErrMsg = kBas_ErrMsg
        lb_ErrMsg.Visible = True
        Timer2.Enabled = True
    End If
    PNL_Start.Visible = False
End Sub
'+--------------------------------------+
'+                                      +
'+      かんばん読取後処理 終了          +
'+                                      +
'+--------------------------------------+
Private Function KanbanProc()
    Dim wKey        As String
    kBas_Poket = 0
    kBas_ColSu = 2  ' 1ﾎﾟｹｯﾄのｶﾗﾑ数
' 読取ｴﾗｰ･自社外かんばん･ﾁｪｯｸﾃﾞｨｼﾞｯﾄｴﾗｰ
    If QRArea.ステータス = 1 Or QRArea.ステータス = 2 Or QRArea.ステータス = 3 Then
       kBas_Poket = kBas_PokErr
       GoTo KanbanProc_Set:
    End If

    Select Case mShiwake
        Case 1
            If QRArea.品番 = "" Then Exit Function
            wKey = Trim(QRArea.品番) '& Trim(QRArea.取引先) & Trim(QRArea.受入)
        Case 2
            If QRArea.背番号 = "" Then Exit Function
            wKey = QRArea.背番号
        Case Else
            If QRArea.取引先 = "" Then Exit Function
            wKey = QRArea.取引先
    End Select
        
    ' 自社外かんばん
    If QRArea.発行元 <> Qr_HMoto Or _
       QRArea.自社コード <> Qr_Jisya Then
       kBas_Poket = kBas_PokErr
       GoTo KanbanProc_Set:
    End If
    ' 対象外かんばん(受注かんばん以外）
    If QRArea.品目 <> 0 Then
       kBas_Poket = kBas_PokErr
       GoTo KanbanProc_Set:
    End If
    
    ' 読取かんばん内容チェック
    With kBas_Check
         .HCKB = H1Op_JKbn(1).Tag
         .HYMD = iH1_Ymd
         .HBIN = iH1_Bin
    End With
    If Not cKanbanChkJ Then
       KanbanProc = False
       GoTo KanbanProc_Set
    End If
    If kBas_Poket <> 0 And kBas_Poket < 5 Then GoTo KanbanProc_Set
    If kBas_Poket <> 0 Then GoTo KanbanProc_Set
    kBas_Poket = 0
    
    
    For kBas_PCnt = 6 To kBas_PoketMax
        If Trim(kBas_TKey(kBas_PCnt)) = "" Or _
           Trim(kBas_TKey(kBas_PCnt)) = Trim(wKey) Then
           ' ﾎﾟｹｯﾄの段数を求める
           kBas_Row = kBas_PCnt Mod 5
           If kBas_Row = 0 Then
               kBas_Row = 5
           End If
           ' 指定ポケット先枚数が最大値を超えないか確認し
           ' ﾎﾟｹｯﾄ最大数がZEROの時は、最大値ﾁｪｯｸをしない
           If kBas_TMai(kBas_PCnt) < kBas_TMaiMax(kBas_Row) Or _
              kBas_TMaiMax(kBas_Row) = 0 Then
              kBas_TKey(kBas_PCnt) = Trim(wKey)
              kBas_Poket = kBas_PCnt
              Exit For
           End If
        End If
    Next kBas_PCnt
    ' ﾎﾟｹｯﾄ指定先がない時は、ﾎﾟｹｯﾄｵｰﾊﾞｰへ
    If kBas_Poket = 0 Then
       kBas_Poket = 4
    End If
KanbanProc_Set:
' ｸﾞﾘｯﾄﾞへ内容セット & 枚数カウント
    kBas_TMai(kBas_Poket) = kBas_TMai(kBas_Poket) + 1
    
    Call GridColor
    
    If GridIDGet Then
       With vsGrid
            If Trim(.TextMatrix(kBas_Row, kBas_Col)) = "" Then
               If mShiwake = 0 Then     ' 納入先仕分け
                  If TorNmGet(kBas_TKey(kBas_Poket), 9) Then
                     .TextMatrix(kBas_Row, kBas_Col) = RTrim(TRM_RName)
                  Else
                     .TextMatrix(kBas_Row, kBas_Col) = kBas_TKey(kBas_Poket)
                  End If
               End If
               If mShiwake = 1 Then     ' 品番仕分け(表示は背番号)
'                  .TextMatrix(kBas_Row, kBas_Col) = Mid(kBas_TKey(kBas_Poket), 1, 11) & vbCr & _
'                                                    Mid(kBas_TKey(kBas_Poket), 12)
                  .TextMatrix(kBas_Row, kBas_Col) = QRArea.背番号
               End If
               If mShiwake = 2 Then     ' 背番号仕分け
                  .TextMatrix(kBas_Row, kBas_Col) = kBas_TKey(kBas_Poket)
               End If
            End If
            .TextMatrix(kBas_Row, kBas_Col + 1) = Format(kBas_TMai(kBas_Poket), "###")
            .Row = kBas_Row: .Col = kBas_Col + 1
            .CellBackColor = mNColor
       End With
    End If

KanbanProc_ED:
    kBas_ReadFlg = True
    ' 枚数表示
    lb_InCnt = CisFun.Val2(lb_InCnt) + 1
    If kBas_Poket < 4 Then
        lb_ErrCnt = Format(CisFun.Val2(lb_ErrCnt) + 1, "#,###")
    Else
        lb_OKCnt = Format(CisFun.Val2(lb_OKCnt) + 1, "#,###")
    End If
End Function
'+------------------------------------+
'+      グリッド位置(Row,Col)を求める
'+------------------------------------+
Private Function GridIDGet() As Boolean
    GridIDGet = False
    ' かんばんポケット最大数を超えた時、ポケットオーバー
    If kBas_Poket > kBas_PoketMax Then
        kBas_Col = 1: kBas_Row = 4
        GridIDGet = True
        Exit Function
    End If
    
    kBas_Col = kBas_Poket \ 5
    kBas_Row = kBas_Poket Mod 5
    If kBas_Row = 0 Then
        kBas_Row = 5
        kBas_Col = kBas_Col - 1
        If kBas_Col < 0 Then kBas_Col = 0
    End If
    kBas_Col = kBas_Col * kBas_ColSu + 1
    
    If kBas_Poket > 30 Then
        kBas_Row = kBas_Row + 6
        kBas_Col = kBas_Col - 12
    End If
    
    GridIDGet = True
End Function
'**************************************
'*          グリッド　初期化           *
'**************************************
Private Sub GridClear()
    With vsGrid
        .Redraw = False
        For mRId = 0 To .Rows - 1
            For mCId = 0 To .Cols - 1
                .TextMatrix(mRId, mCId) = ""
                If mCId Mod 2 = 1 Then          ' 仕入先情報(左詰)
                   .Row = mRId: .Col = mCId
                   .CellAlignment = flexAlignLeftCenter
                Else                            ' 枚数情報(右詰)
                   .Row = mRId: .Col = mCId
                   .CellAlignment = flexAlignRightCenter
                   .CellBackColor = &HFFFFC0
                End If
            Next mCId
        Next mRId
        .Redraw = True
    End With
End Sub
Private Sub GridColor()
    For gCnt1 = 1 To 11
        With vsGrid
            .Redraw = False
            For gCnt2 = 1 To 6
                .Row = gCnt1: .Col = gCnt2 * 2
                .CellBackColor = mMColor
            Next gCnt2
            .Redraw = True
        End With
    Next gCnt1
End Sub
'+---------------------------------------+
'+      名称マスタ初期設定
'+---------------------------------------+
Private Sub MeisyoInit()
'   <><><><><> 受注区分 <><><><><>
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "受注区分"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "通常"
        .作成者 = gTanto
        .制御区分 = "X"
    End With
    If Not SYMInitCreate(True) Then
        With SYMCreate
            .値 = "2":  .値名称 = "特配":       Call SYMInitCreate
            .値 = "3":  .値名称 = "臨時":       Call SYMInitCreate
            .値 = "4":  .値名称 = "スポット":   Call SYMInitCreate
            .値 = "5":  .値名称 = "指示":       Call SYMInitCreate
            .値 = "6":  .値名称 = "号試":       Call SYMInitCreate
            .値 = "7":  .値名称 = "代替":       Call SYMInitCreate
            .値 = "8":  .値名称 = "その他":     Call SYMInitCreate
        End With
    End If
'   >> 受注区分の表示
    gSL_Select = ""
    gSL_Select = gSL_Select & "select 値"
    gSL_Select = gSL_Select & "     , 値名称"
    gSL_Select = gSL_Select & "  from 名称マスタ"
    gSL_Select = gSL_Select & " where 区分名称 = '受注区分'"
    gSL_Select = gSL_Select & "   and 区分タイプ = ''"
    gSL_Select = gSL_Select & " order by 値"
    If SYMRead(gSL_Select, 0) Then
        Do Until Not SYM_RDSTS
            Select Case SYM.値
                Case "1":   H1Op_JKbn(1).Caption = SYM.値名称:  H1Op_JKbn(1).Tag = SYM.値
                Case "5":   H1Op_JKbn(2).Caption = SYM.値名称:  H1Op_JKbn(2).Tag = SYM.値
                Case "6":   H1Op_JKbn(3).Caption = SYM.値名称:  H1Op_JKbn(3).Tag = SYM.値
                Case "4":   H1Op_JKbn(4).Caption = SYM.値名称:  H1Op_JKbn(4).Tag = SYM.値
                Case "3":   H1Op_JKbn(5).Caption = SYM.値名称:  H1Op_JKbn(5).Tag = SYM.値
                Case "2":   H1Op_JKbn(6).Caption = SYM.値名称:  H1Op_JKbn(6).Tag = SYM.値
                Case "7":   H1Op_JKbn(7).Caption = SYM.値名称:  H1Op_JKbn(7).Tag = SYM.値
                Case "8":   H1Op_JKbn(8).Caption = SYM.値名称:  H1Op_JKbn(8).Tag = SYM.値
            End Select
            Call SYMReadNext
        Loop
    End If
    Call SYMClose
End Sub
'+---------------------------------------+
'+      終了処理
'+---------------------------------------+
Private Sub End_Proc()
    ' かんばんエラーリスト
    gSL_Select = "SELECT count(*) 件数 FROM 帳票出力テーブル CH "
    gSL_Select = gSL_Select & "  Left Outer Join かんばん読取エラーテーブル KE"
    gSL_Select = gSL_Select & "      ON  KE.エラー管理番号 = CH.管理NO"
    gSL_Select = gSL_Select & "      And KE.エラー種類 = CH.帳票区分"
    gSL_Select = gSL_Select & " Where 端末番号 = Host_Name()"
    gSL_Select = gSL_Select & "   And 帳票区分 = 41 "
    gSL_Select = gSL_Select & "   And ISNULL(KE.品番,'') <> ''"
    Call CHTRead(gSL_Select, 1)
    If CHT.件数 <> 0 And kBas_Check.OpJErrPrt = 0 Then                      '2005/11/08条件追加
        If Not CXJ0031LPrint Then Exit Sub
    End If
    
    gSL_Select = "SELECT count(*) 件数 FROM 受注かんばん読取テーブル "
    Call CHTRead(gSL_Select, 1)
    If CHT.件数 = 0 Then Exit Sub
    With CisDB
         .SQL = "INSERT INTO 受注入力ワーク" & vbCr
         .SQL = .SQL & "(端末番号,受注管理NO,受注管理行NO, 受注区分, 受注日, 受注便, 納入日," & vbCr
         .SQL = .SQL & "         納入便, 納入先, 出荷予定日, 出荷予定便, 受入, 機番," & vbCr
         .SQL = .SQL & "         納入番号, 納品書頁, 納品書行, 分納区分, 分納回数," & vbCr
         .SQL = .SQL & "         品番, 背番号, 枚数, 収容数, 納入数, 端数," & vbCr
         .SQL = .SQL & "         出荷指示書NO, 出荷指示書行NO, 抜取枚数,納品書発行区分," & vbCr
         .SQL = .SQL & "         エフ発行枚数, エフ発行端数," & vbCr
         .SQL = .SQL & "         作成種別, 累積区分, 納品書管理NO, 元出荷指示書NO," & vbCr
         .SQL = .SQL & "         元出荷指示書行NO)" & vbCr
         .SQL = .SQL & " Select " & vbCr
         .SQL = .SQL & " Host_Name(),受注管理NO,1,受注区分,Convert(Varchar(8),Getdate(),112),''," & vbCr
         .SQL = .SQL & " 納入日,納入便,納入先,'','',受入,''," & vbCr
         .SQL = .SQL & " '',0,0,0,0," & vbCr
         .SQL = .SQL & " 品番,背番号,枚数,収容数,受注数,0," & vbCr
         .SQL = .SQL & " 0,0,0,0, " & vbCr
         .SQL = .SQL & " 0,0," & vbCr
         .SQL = .SQL & " 'K',0,0,0," & vbCr
         .SQL = .SQL & " 0 " & vbCr
         .SQL = .SQL & " From 受注かんばん読取テーブル "
         If Not .DBExec Then Exit Sub
         
         If .RecordCount <> 0 Then
         '   出荷指示NO採番
'*--------------------------------------------------<< 2004/06/24 Insert Start >>
             gCHT_DelKbn53 = 0
'*--------------------------------------------------<< 2004/06/24 Insert End >>
             CXJ0110.Show vbModal
             Unload CXJ0110
             Set CXJ0110 = Nothing
             DoEvents
         '   出荷指示書発行
             Call CXJ0120LPrint(0)
         '   自社伝票発行
             Call CXJ0140LPrint(0)
         '   自社伝票ﾌｧｲﾙ出力
             Call CXJ0150TFile(0)
         '   ノーかんばんエフ発行
'                Call CXJ0130LPrint(0)
         End If
         
    End With
End Sub
