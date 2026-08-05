VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{490AF35D-478A-11D4-AEC2-00000EA57F9E}#4.0#0"; "CISMSG~2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Begin VB.Form CKN0010 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "かんばん仕分"
   ClientHeight    =   11010
   ClientLeft      =   345
   ClientTop       =   855
   ClientWidth     =   14565
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
   ScaleHeight     =   11010
   ScaleWidth      =   14565
   WindowState     =   2  '最大化
   Begin VB.PictureBox H1_Area1 
      Height          =   735
      Left            =   1200
      ScaleHeight     =   675
      ScaleWidth      =   4455
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1905
      Width           =   4515
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   585
         Left            =   60
         Top             =   45
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1032
         ForeColor       =   16711680
         Caption         =   "仕分け方法"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   60
      End
      Begin Cis3D_v60.CIS3D OpH1_Group 
         Height          =   585
         Index           =   0
         Left            =   1890
         Top             =   45
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1032
         BackColor       =   16776960
         ForeColor       =   0
         Caption         =   "仕入先"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderWidth     =   5
         cAlingnment     =   7
         cButton         =   -1  'True
         cPositionY      =   30
      End
      Begin Cis3D_v60.CIS3D OpH1_Group 
         Height          =   585
         Index           =   1
         Left            =   3165
         Top             =   45
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1032
         ForeColor       =   0
         Caption         =   "品番"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderWidth     =   5
         cAlingnment     =   7
         cButton         =   -1  'True
         cPositionY      =   30
      End
   End
   Begin Cis3D_v60.CIS3D PNL_PF 
      Height          =   855
      Left            =   8355
      Top             =   2025
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
      Left            =   1440
      Top             =   1620
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
      TabIndex        =   3
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
      TabIndex        =   1
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
               Left            =   450
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
               Left            =   450
               Top             =   1110
               Width           =   1395
               _ExtentX        =   2461
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
               TabIndex        =   2
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
      Caption         =   "【 かんばん仕分 】"
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
      Left            =   1185
      Top             =   2985
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
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   885
      Width           =   225
   End
End
Attribute VB_Name = "CKN0010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   かんばん仕分 [仕入先] Or [品番]
'**       フォームID    :   CKN0010
'**       処理概要      :
'**
'**       作  成  日    :   2004/01/16  By CIS
'**       変  更  日    :
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
    
    mGotColor = OpH1_Group(0).BackColor
    mLostColor = OpH1_Group(1).BackColor
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
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
    
' かんばん処理 初期設定
    Set TX597G_Comm = MSComm1
    Call TX597GInfoGet
    kBas_Proc = "S"
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
    
    Call OpH1_Group_Click(0)
    
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
           Case vbKeyF1:        If ProcHB = "H1" Then Call OpH1_Group_Click(0)  '【仕入先】
           Case vbKeyF2:        If ProcHB = "H1" Then Call OpH1_Group_Click(1)  '【品番】
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                      '【入力】
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
'+----------------------------+
'+  仕分け方法選択             +
'+----------------------------+
Private Sub OpH1_Group_Click(Index As Integer)
    If Not OpH1_Group(Index).Enabled Then Exit Sub
    If Trim(OpH1_Group(0).Tag) = "" Then OpH1_Group(0).Tag = "99"
    If Index = CisFun.Val2(OpH1_Group(0).Tag) Then Exit Sub
    
    If PB_ENT2.Visible Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     読取枚数の初期化を行います！      "
            .MB_MSG(4) = "     　　　よろしいですか？           "
            .MB_Title = "仕分方法変更"
            .MB_Button = OK_CAN
            If Not .MBOX Then Exit Sub
       End With
       Call PB_ENT2_Click
    End If
    For Each gObj In OpH1_Group
        gObj.BackColor = mLostColor
        gObj.cButton = True
    Next gObj
    
    OpH1_Group(Index).BackColor = mGotColor
    OpH1_Group(Index).cButton = False
    OpH1_Group(Index).cBoderStyle = 1
    OpH1_Group(0).Tag = Format(Index, "0")

    mShiwake = Index

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

    Call KanbanRead
    
    ProcHB = "H1"
    Call DispChange(ProcHB)
    
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
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
            gStr = Trim(QRArea.品番) & Trim(QRArea.取引先) & Trim(QRArea.受入)
        Case 2
            If QRArea.背番号 = "" Then Exit Function
            gStr = QRArea.背番号
        Case Else
            If QRArea.取引先 = "" Then Exit Function
            gStr = QRArea.取引先
    End Select
        
    
    For kBas_PCnt = 3 To kBas_PoketMax
        If Trim(kBas_TKey(kBas_PCnt)) = "" Or _
           Trim(kBas_TKey(kBas_PCnt)) = Trim(gStr) Then
           ' ﾎﾟｹｯﾄの段数を求める
           kBas_Row = kBas_PCnt Mod 5
           If kBas_Row = 0 Then
               kBas_Row = 5
           End If
           ' 指定ポケット先枚数が最大値を超えないか確認し
           ' ﾎﾟｹｯﾄ最大数がZEROの時は、最大値ﾁｪｯｸをしない
           If kBas_TMai(kBas_PCnt) < kBas_TMaiMax(kBas_Row) Or _
              kBas_TMaiMax(kBas_Row) = 0 Then
              kBas_TKey(kBas_PCnt) = Trim(gStr)
              kBas_Poket = kBas_PCnt
              Exit For
           End If
        End If
    Next kBas_PCnt
    ' ﾎﾟｹｯﾄ指定先がない時は、ﾎﾟｹｯﾄｵｰﾊﾞｰへ
    If kBas_Poket = 0 Then
       kBas_Poket = 2
    End If
KanbanProc_Set:
' ｸﾞﾘｯﾄﾞへ内容セット & 枚数カウント
    kBas_TMai(kBas_Poket) = kBas_TMai(kBas_Poket) + 1
    If GridIDGet Then
       With vsGrid
            If Trim(.TextMatrix(kBas_Row, kBas_Col)) = "" Then
               If mShiwake = 0 Then     ' 仕入先仕分け
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
       End With
    End If

KanbanProc_ED:
    kBas_ReadFlg = True
    ' 枚数表示
    lb_InCnt = CisFun.Val2(lb_InCnt) + 1
    If kBas_Poket < 3 Then
        lb_ErrCnt = Format(CisFun.Val2(lb_ErrCnt) + 1, "#,###")
    Else
        lb_OKCnt = Format(CisFun.Val2(lb_OKCnt) + 1, "#,###")
    End If
End Function
'+------------------------------------+
'+      グリッド位置(Row,Col)を求める
'+------------------------------------+
'Private Function GridIDGet() As Boolean                                      ' 2006/02/03 Delete
Private Function GridIDGet(Optional MaxCheck As Boolean = True) As Boolean    ' 2006/02/03 Insert
    
    kBas_ColSu = 2  ' 1ﾎﾟｹｯﾄのｶﾗﾑ数

    GridIDGet = False
'---------------------------------------------------------------------------- 2006/02/03 Delete
'    ' かんばんポケット最大数を超えた時、ポケットオーバー
'    If kBas_Poket > kBas_PoketMax Then
'        kBas_Col = 1: kBas_Row = 2
'        GridIDGet = True
'        Exit Function
'    End If
'---------------------------------------------------------------------------- 2006/02/03 Delete End
        
'---------------------------------------------------------------------------- 2006/02/03 Insert
    If MaxCheck Then
       If kBas_Poket > kBas_PoketMax Then
          kBas_Col = 1: kBas_Row = 2
          GridIDGet = True
          Exit Function
       End If
    End If
'---------------------------------------------------------------------------- 2006/02/03 Insert End
    
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
'--------------------------------------------------------- 2006/02/03 Insert
    Call PoketMaxColorSet
'--------------------------------------------------------- 2006/02/03 Insert End
End Sub
'--------------------------------------------------------- 2006/02/03 Insert
'+---------------------------------------+
'+      ポケット最大数を超えた時色セット
'+---------------------------------------+
Private Sub PoketMaxColorSet()
    

    If kBas_PoketMax < 61 And kBas_MaxPColor <> 0 Then
       kBas_PoketSave = kBas_Poket
       With vsGrid
            .Redraw = False
            
            For gCnt1 = kBas_PoketMax + 1 To 60
                kBas_Poket = gCnt1
                Call GridIDGet(False)
                .Row = kBas_Row: .Col = kBas_Col
                .CellBackColor = kBas_MaxPColor
                .Col = kBas_Col + 1: .CellBackColor = kBas_MaxPColor
            Next gCnt1
            .Redraw = True
       End With
       kBas_Poket = kBas_PoketSave
    End If
End Sub
'--------------------------------------------------------- 2006/02/03 Insert End

