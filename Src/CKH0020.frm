VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{490AF35D-478A-11D4-AEC2-00000EA57F9E}#4.0#0"; "CISMSG~2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Begin VB.Form CKH0020 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "発注かんばん入力 [一括] 仕入先単位"
   ClientHeight    =   10530
   ClientLeft      =   210
   ClientTop       =   570
   ClientWidth     =   14385
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
   ScaleHeight     =   10530
   ScaleWidth      =   14385
   WindowState     =   2  '最大化
   Begin VB.PictureBox PNL_Start 
      BackColor       =   &H00000000&
      Height          =   825
      Left            =   4605
      ScaleHeight     =   765
      ScaleWidth      =   7125
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2430
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
      Height          =   5085
      Left            =   300
      ScaleHeight     =   5025
      ScaleWidth      =   14550
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4200
      Width           =   14610
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   4920
         Left            =   60
         Top             =   60
         Width           =   14460
         _ExtentX        =   25506
         _ExtentY        =   8678
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
            Height          =   4815
            Index           =   1
            Left            =   45
            Top             =   45
            Width           =   14370
            _ExtentX        =   25347
            _ExtentY        =   8493
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
               Height          =   585
               Index           =   4
               Left            =   12150
               Top             =   1095
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   1032
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
               Height          =   585
               Index           =   5
               Left            =   13200
               Top             =   1095
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   1032
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
               Height          =   585
               Index           =   0
               Left            =   7950
               Top             =   1095
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   1032
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
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   585
               Index           =   1
               Left            =   9000
               Top             =   1095
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   1032
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
               Height          =   585
               Index           =   2
               Left            =   10050
               Top             =   1095
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   1032
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
               Height          =   585
               Index           =   3
               Left            =   11100
               Top             =   1095
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   1032
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
               Height          =   585
               Index           =   29
               Left            =   120
               Top             =   1095
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   1032
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
               Height          =   585
               Index           =   30
               Left            =   450
               Top             =   1095
               Width           =   1650
               _ExtentX        =   2910
               _ExtentY        =   1032
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
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   585
               Index           =   31
               Left            =   2100
               Top             =   1095
               Width           =   1650
               _ExtentX        =   2910
               _ExtentY        =   1032
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
               Height          =   585
               Index           =   32
               Left            =   3750
               Top             =   1095
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   1032
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
               Height          =   585
               Index           =   33
               Left            =   4800
               Top             =   1095
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   1032
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
               Height          =   585
               Index           =   34
               Left            =   5850
               Top             =   1095
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   1032
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
               Height          =   585
               Index           =   35
               Left            =   6900
               Top             =   1095
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   1032
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
               Height          =   600
               Index           =   36
               Left            =   120
               Top             =   1680
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   1058
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
               Height          =   600
               Index           =   1
               Left            =   450
               Top             =   1680
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   1058
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
               Height          =   600
               Index           =   37
               Left            =   120
               Top             =   2280
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   1058
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
               Height          =   600
               Index           =   2
               Left            =   450
               Top             =   2280
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   1058
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
               Height          =   600
               Index           =   38
               Left            =   120
               Top             =   2880
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   1058
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
            Begin Cis3D_v60.CIS3D B1_ErrName 
               Height          =   600
               Index           =   3
               Left            =   450
               Top             =   2880
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   1058
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
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   600
               Index           =   39
               Left            =   120
               Top             =   3480
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   1058
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
            Begin Cis3D_v60.CIS3D B1_ErrName 
               Height          =   600
               Index           =   4
               Left            =   450
               Top             =   3480
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   1058
               BackColor       =   16761087
               ForeColor       =   0
               Caption         =   "抜き取り"
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
               Height          =   600
               Index           =   40
               Left            =   120
               Top             =   4080
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   1058
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
            Begin Cis3D_v60.CIS3D B1_ErrName 
               Height          =   600
               Index           =   5
               Left            =   450
               Top             =   4080
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   1058
               BackColor       =   16761087
               ForeColor       =   0
               Caption         =   "あふれ"
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
               Height          =   600
               Index           =   6
               Left            =   2100
               Top             =   1680
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   1058
               BackColor       =   16761087
               ForeColor       =   0
               Caption         =   "打切(完了)"
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
               Height          =   600
               Index           =   7
               Left            =   2100
               Top             =   2280
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   1058
               BackColor       =   16761087
               ForeColor       =   0
               Caption         =   "打切(端数)"
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
               Height          =   600
               Index           =   8
               Left            =   2100
               Top             =   2880
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   1058
               BackColor       =   16761087
               ForeColor       =   0
               Caption         =   "平準化"
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
               Height          =   600
               Index           =   9
               Left            =   2100
               Top             =   3480
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   1058
               BackColor       =   16761087
               ForeColor       =   0
               Caption         =   "ﾎﾟｹｯﾄｵｰﾊﾞｰ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
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
               Height          =   600
               Index           =   10
               Left            =   2100
               Top             =   4080
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   1058
               BackColor       =   16761087
               ForeColor       =   0
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
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   1
               cPositionX      =   30
            End
            Begin vsFlexLib.vsFlexArray vsGrid1 
               Height          =   3660
               Left            =   90
               TabIndex        =   2
               TabStop         =   0   'False
               Top             =   1050
               Width           =   14190
               _Version        =   196608
               _ExtentX        =   25030
               _ExtentY        =   6456
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
               FormatString    =   "№|XXXXXXXXX1XX3|XXXXXXXXX1XX3|XXXXXXX8|XXXXXXX8|XXXXXXX8|XXXXXXX8|XXXXXXX8|XXXXXXX8|XXXXXXX8|XXXXXXX8|XXXXXXX8|XXXXXXX8|"
               Rows            =   6
               Cols            =   20
               BackColor       =   16777152
               BackColorBkg    =   12632256
               FocusRect       =   0
               HighLight       =   0
               ScrollBars      =   0
               RowHeightMin    =   600
               GridLines       =   2
               SelectionMode   =   1
            End
            Begin vsFlexLib.vsFlexArray vsGrid2 
               Height          =   915
               Left            =   3720
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   135
               Width           =   10560
               _Version        =   196608
               _ExtentX        =   18627
               _ExtentY        =   1614
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
               FormatString    =   "XXXXXXX8|XXXXXXX8|XXXXXXX8|XXXXXXX8|XXXXXXX8|XXXXXXX8|XXXXXXX8|XXXXXXX8|XXXXXXX8|XXXXXXX8|"
               Rows            =   3
               Cols            =   20
               FixedCols       =   0
               BackColor       =   14737632
               BackColorFixed  =   14737632
               BackColorBkg    =   12632256
               FocusRect       =   0
               HighLight       =   0
               ScrollBars      =   0
               RowHeightMin    =   280
               GridLines       =   9
               GridLinesFixed  =   9
               SelectionMode   =   1
            End
            Begin Cis3D_v60.CIS3D B1_Name 
               Height          =   915
               Left            =   2070
               Top             =   150
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   1614
               BackColor       =   14737632
               Caption         =   "XXXXXXX8"
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
               cAlingnment     =   6
               cBoderStyle     =   1
               cPositionY      =   30
               Begin Cis3D_v60.CIS3D B1_Name2 
                  Height          =   345
                  Index           =   0
                  Left            =   120
                  Top             =   300
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   609
                  BackColor       =   14737632
                  Caption         =   "XXXXXXXXX1XX"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cBoderColor1    =   14737632
                  cBoderColor2    =   14737632
                  cFont3DColor1   =   8421504
                  cFont3DColor2   =   -2147483643
                  cFont3D         =   2
                  cBoderStyle     =   2
               End
               Begin Cis3D_v60.CIS3D B1_Name2 
                  Height          =   345
                  Index           =   1
                  Left            =   120
                  Top             =   570
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   609
                  BackColor       =   14737632
                  Caption         =   "XXXXXXXXX1XX"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cBoderColor1    =   14737632
                  cBoderColor2    =   14737632
                  cFont3DColor1   =   8421504
                  cFont3DColor2   =   -2147483643
                  cFont3D         =   2
                  cBoderStyle     =   2
               End
            End
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
      TabIndex        =   0
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
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 発注かんばん入力　[一括] 仕入先単位 】"
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
         Left            =   14265
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
         Left            =   11745
         Top             =   30
         Width           =   2505
         _ExtentX        =   4419
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
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   420
         Left            =   0
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
   End
   Begin Cis3D_v60.CIS3D PNL_Kensu 
      Height          =   630
      Left            =   345
      Top             =   3435
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
      Left            =   1980
      Top             =   2370
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
   Begin Cis3D_v60.CIS3D PNL_PF 
      Height          =   855
      Left            =   9330
      Top             =   1320
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
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin Cis3D_v60.CIS3D PoketErrMsg 
      Height          =   585
      Left            =   390
      Top             =   9630
      Visible         =   0   'False
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   1032
      BackColor       =   16761087
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
      cAlingnment     =   7
   End
End
Attribute VB_Name = "CKH0020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   発注かんばん入力[一括]仕入先単位
'**       フォームID    :   CKH0020
'**       処理概要      :
'**
'**       作  成  日    :   2004/01/20  By CIS SIMPR-A
'**       変  更  日    :   2006/02/03 By CIS ポケット最大数を越えたポケットの色設定
'**                                           エラーポケット可変対応
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid1      As New CisVsGrid3
    Dim CisVsGrid2      As New CisVsGrid3
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim mSaveRow       As Integer       ' Row退避
    Dim mSaveCol       As Integer       ' Col退避
    
    Dim mKey            As String

    Dim mRId            As Long
    Dim mCId            As Long

    Dim mMColor         As Long
    Dim mNColor         As Long

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
        
    mMColor = &HFFFFC0
    mNColor = &HFFC0C0

    '+-----------------------+
    '+  グリッド初期設定(枚数)
    '+-----------------------+
    With CisVsGrid1
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 600
        .PatanMax = 1
        .SelectCol = 19
        .InitGet = False
        .Init
    End With
    Call GridInit1
    '+----------------------------+
    '+  グリッド初期設定(仕入先)
    '+----------------------------+
    With CisVsGrid2
        Set .GridObj = vsGrid2
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 280
        .PatanMax = 1
        .SelectCol = 10
        .InitGet = False
        .Init
    End With
    Call GridInit2
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "B1"
    Call DispChange(ProcHB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

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
    
    If Not cKanbanInit Then
        Call PB_END_Click
    End If
'----------------------------------------------------------- 2006/02/03 Delete
''初期処理・初期表示
'    If Not Init_Set Then
'        Call PB_END_Click
'    End If
'----------------------------------------------------------- 2006/02/03 Delete End
'----------------------------------------------------------- 2006/02/03 Insert
'初期処理・初期表示
    If Not GridErrSet Then
        Call PB_END_Click
    End If
    Call GridColor
'----------------------------------------------------------- 2006/02/03 Insert End
    
    If kBas_ExecMode Then
       K_Sykbnm.Visible = True
    Else
       K_Sykbnm.Visible = False
    End If
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
           Case vbKeyS:        If kBas_ExecMode Then kBas_DemoFlg = "S"
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
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
'+   後方処理
'+---------------------+
    If kBas_Check.ErrCnt <> 0 Or kBas_Check.OutCnt <> 0 Then
       CKH0900.Show vbModal
       Unload CKH0900
       Set CKH0900 = Nothing
       DoEvents
        ' 発注処理対象データ確認
       gSL_Select = "Select Count(*) 件数 From 発注かんばん予定テーブル "
       gSL_Select = gSL_Select & " Where 処理端末 = Host_Name()"
       gSL_Select = gSL_Select & "   And 処理区分 = 1"
       Call IRNRead(gSL_Select, 1)
       If IRN.件数 <> 0 Then
        ' 発注処理
           CisDB.PassLink
           On Error Resume Next
           With CisFun
                .ShellApl = gIniExe & "PXH0110.exe"
                .ShellAplPara = "K"
                .ShellTask NoFreeze
           End With
           On Error GoTo 0
        End If
    End If
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
   Call KanbanRead
End Sub
'+----------------------------+
'+    枚数初期化(F1)           +
'+----------------------------+
Private Sub PB_ENT2_Click()
    
    Erase kBas_TKey
    Erase kBas_TKey2
    
    Erase kBas_TMai
    
    ReDim kBas_TKey(kBas_PoketMax)
    ReDim kBas_TKey2(kBas_PoketMax)
    ReDim kBas_TMai(kBas_PoketMax)
    
    Call GridClear
    
    kBas_ReadFlg = False
    PB_ENT2.Visible = False
    
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
     Case "RD"
          SyoriNM = "読取"
          K_Sykbnm.ForeColor = &HFFFF&         '(黄色)
     Case Else
          SyoriNM = ""
   End Select
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
'****************************
'*                          *
'*      HEAD/BODY 処理      *
'*                          *
'****************************
'****************************
'*      入力切替 処理        *
'****************************
Private Function DispChange(Pro As String)
    Dim B1Mode      As Boolean
    Dim B1Color     As Long
    Dim RDMode      As Boolean
    Dim RDColor     As Long
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
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
       Case "RD"
          RDMode = True
          RDColor = gPGotSel
    End Select
    
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    PNL_PF.Visible = B1Mode
    PB_ENT2.Visible = kBas_ReadFlg
    
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
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
End Sub
'*************************************************************************
'*      入力処理(Enter)
'*************************************************************************
Private Sub ReturnPress()
'    Call CisFrm.MousePT(11)
'    If ProcHB = "B1" Then
'ReturnPress_Ed:
'   Call CisFrm.MousePT(1)
End Sub
'*------------------------------------------------------------------------'
'*          Head1 チェック
'*------------------------------------------------------------------------'
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
    Head1Chk = True
End Function
'-------------------------------------------------------------------- 2006/02/03 Delete
''+---------------------------------------+
''+      初期設定・初期表示
''+---------------------------------------+
'Private Function Init_Set() As Boolean
'    Init_Set = False
''<< かんばん読取エラーポケット表示 >>
'    gSL_Select = " SELECT * FROM 名称マスタ"
'    gSL_Select = gSL_Select & " WHERE 区分名称 = 'かんばんエラーポケット'"
'    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
'    gSL_Select = gSL_Select & "   And 数字1 = 0"
'    gSL_Select = gSL_Select & " ORDER BY 値"
'    If SYMRead(gSL_Select, 0) Then
'        gInt = 0
'        Do Until Not SYM_RDSTS
'            gInt = gInt + 1
'            B1_ErrName(gInt) = RTrim(SYM.値名称)
'            Call SYMReadNext
'        Loop
'    End If
'    Call SYMClose
'
'    Init_Set = True
'End Function
'-------------------------------------------------------------------- 2006/02/03 Delete End
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit1()
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 2, 0, "NO")
        Call .FixedSet(中中, 右中, あり, 13, 1, "枚数1")
        Call .FixedSet(中中, 右中, あり, 13, 2, "枚数2")
        Call .FixedSet(中中, 右中, あり, 8, 3, "枚数3")
        Call .FixedSet(中中, 右中, あり, 8, 4, "枚数4")
        Call .FixedSet(中中, 右中, あり, 8, 5, "枚数5")
        Call .FixedSet(中中, 右中, あり, 8, 6, "枚数6")
        Call .FixedSet(中中, 右中, あり, 8, 7, "枚数7")
        Call .FixedSet(中中, 右中, あり, 8, 8, "枚数8")
        Call .FixedSet(中中, 右中, あり, 8, 9, "枚数9")
        Call .FixedSet(中中, 右中, あり, 8, 10, "枚数10")
        Call .FixedSet(中中, 右中, あり, 8, 11, "枚数11")
        Call .FixedSet(中中, 右中, あり, 8, 12, "枚数12")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
'        Call .FixedSet(中中, 左中, なし, 0, 19, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12")
        .InitDisp
    End With
    
End Sub
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit2()
    With CisVsGrid2
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 左中, あり, 8, 0, "")
        Call .FixedSet(中中, 左中, あり, 8, 1, "")
        Call .FixedSet(中中, 左中, あり, 8, 2, "")
        Call .FixedSet(中中, 左中, あり, 8, 3, "")
        Call .FixedSet(中中, 左中, あり, 8, 4, "")
        Call .FixedSet(中中, 左中, あり, 8, 5, "")
        Call .FixedSet(中中, 左中, あり, 8, 6, "")
        Call .FixedSet(中中, 左中, あり, 8, 7, "")
        Call .FixedSet(中中, 左中, あり, 8, 8, "")
        Call .FixedSet(中中, 左中, あり, 8, 9, "")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
'        Call .FixedSet(中中, 左中, なし, 0, 19, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9")
        .InitDisp
    End With
    
    For gCnt1 = 1 To 2
        With vsGrid2
            .Row = gCnt1
            For gCnt2 = 0 To 9
                .Col = gCnt2
                .CellFontSize = 10
                .Text = ""
            Next gCnt2
        End With
    Next gCnt1
'-------------------------------------------- 2006/02/03 Insert
    B1_Name = ""
    B1_Name2(0) = ""
    B1_Name2(1) = ""
'-------------------------------------------- 2006/02/03 Insert End
    
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
    ProcHB = "B1"
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
'    ProcMsg.MsgActED
    ProcMsg = "【 かんばん読取中 】"
    DoEvents
' かんばんソータｽﾃｰﾀｽﾁｪｯｸ
    kBas_ErrSts = TX597GCheck
    
    Select Case kBas_ErrSts
       Case 0
            If Not KanbanProc Then         ' かんばん情報 処理 (設定値 = 指定ポケット№)
               Call ErrDisp("E")
               ReadNext = False
               Exit Function
            End If
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
            lb_ErrMsg.cAlingnment = 6
        Else
            lb_ErrMsg.cPositionY = 0
            lb_ErrMsg.cAlingnment = 7
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
Private Function KanbanProc() As Boolean
    KanbanProc = True
    
    Dim mPSetFlg        As Boolean
    Dim mPoketFlg       As Boolean
    
    kBas_Poket = 0
    ' 読取ｴﾗｰ･自社外かんばん･ﾁｪｯｸﾃﾞｨｼﾞｯﾄｴﾗｰ
    If QRArea.ステータス = 1 Or QRArea.ステータス = 2 Or QRArea.ステータス = 3 Then
       kBas_Poket = kBas_PokErr
       kBas_ErrNo = 99
       GoTo KanbanProc_Set:
    End If
        
    mKey = Trim(QRArea.取引先)
    If Trim(mKey) = "" Then Exit Function
    
    ' 自社外かんばん
    If QRArea.発行元 <> Qr_HMoto Or _
       QRArea.自社コード <> Qr_Jisya Then
       kBas_Poket = kBas_PokErr
       kBas_ErrNo = 99
       GoTo KanbanProc_Set:
    End If
    ' 部品発注かんばん以外はエラー
    If QRArea.品目 <> 1 Then
       kBas_Poket = kBas_PokErr
       kBas_ErrNo = 99
       GoTo KanbanProc_Set:
    End If
    ' 読取かんばん内容チェック
    With kBas_Check
         .ProcKB = "1"
         .HCKB = "1"
         .HYMD = ""
         .HBIN = ""
    End With
    If Not cKanbanChkH Then
       KanbanProc = False
       kBas_ErrNo = 99
       GoTo KanbanProc_Set
    End If
    If kBas_Poket <> 0 Then GoTo KanbanProc_Set
    
    mPoketFlg = False
    
'------------------------------------------------------------- 2006/02/03 Delete
'    For kBas_PCnt = 11 To kBas_PoketMax
'        mPSetFlg = False
'        If Trim(kBas_TKey(kBas_PCnt)) = Trim(mKey) Or _
'           Trim(kBas_TKey(kBas_PCnt)) = "" Then
'           mPSetFlg = True
'        End If
'
'        If mPSetFlg Then
'
'           mPoketFlg = True
'
'           ' ﾎﾟｹｯﾄの段数を求める
'           kBas_Row = kBas_PCnt Mod 5
'           If kBas_Row = 0 Then
'               kBas_Row = 5
'           End If
'           ' 指定ポケット先枚数が最大値を超えないか確認し
'           ' ﾎﾟｹｯﾄ最大数がZEROの時は、最大値ﾁｪｯｸをしない
'           If kBas_TMai(kBas_PCnt) < kBas_TMaiMax(kBas_Row) Or _
'              kBas_TMaiMax(kBas_Row) = 0 Then
'              kBas_Poket = kBas_PCnt
'              Exit For
'           End If
'        End If
'    Next kBas_PCnt
'
'    ' ﾎﾟｹｯﾄ指定先がない時は、ﾎﾟｹｯﾄｵｰﾊﾞｰへｾｯﾄ
'    If kBas_Poket = 0 Then
'       If kBas_TMai(9) < kBas_TMaiMax(5) Or _
'          kBas_TMaiMax(5) = 0 Then
'          kBas_Poket = 9
'       Else
'          kBas_Poket = 10
'       End If
'    End If
'------------------------------------------------------------- 2006/02/03 Delete End
'
'----------------------------------------------------------- 2006/02/03 Insert
    For kBas_PCnt = 1 To kBas_PoketMax
        mPSetFlg = False
        If Trim(kBas_ErSetMsg(kBas_PCnt)) <> "" Then
        Else
            If Trim(kBas_TKey(kBas_PCnt)) = Trim(mKey) Or _
               Trim(kBas_TKey(kBas_PCnt)) = "" Then
               mPSetFlg = True
            End If
        End If
        If mPSetFlg Then
           mPoketFlg = True
           ' ﾎﾟｹｯﾄの段数を求める
           kBas_Row = kBas_PCnt Mod 5
           If kBas_Row = 0 Then
               kBas_Row = 5
           End If
           ' 指定ポケット先枚数が最大値を超えないか確認し
           ' ﾎﾟｹｯﾄ最大数がZEROの時は、最大値ﾁｪｯｸをしない
           If kBas_TMai(kBas_PCnt) < kBas_TMaiMax(kBas_Row) Or _
              kBas_TMaiMax(kBas_Row) = 0 Then
              kBas_Poket = kBas_PCnt
              Exit For
           End If
        End If
    Next kBas_PCnt
    
    ' ﾎﾟｹｯﾄ指定先がない時は、ﾎﾟｹｯﾄｵｰﾊﾞｰへｾｯﾄ
    If kBas_Poket = 0 Then
       kBas_Poket = 9
       kBas_ErrNo = 9
    End If
'----------------------------------------------------------- 2006/02/03 Insert End

KanbanProc_Set:
'----------------------------------------------------------- 2006/02/03 Insert
    If kBas_ErrNo <> 0 Then
        For gLong = 0 To UBound(kBas_ErNo)
            If kBas_ErNo(gLong) = kBas_Poket Then
                If kBas_ErPno(gLong) = 0 Then
                   kBas_Poket = kBas_PoketErr
                Else
                   kBas_Poket = kBas_ErPno(gLong)
                End If
                Exit For
            End If
        Next gLong
    End If
'----------------------------------------------------------- 2006/02/03 Insert End
' ｸﾞﾘｯﾄﾞへ内容セット & 枚数カウント
    kBas_TMai(kBas_Poket) = kBas_TMai(kBas_Poket) + 1
    
    Call GridColor
    
    If GridIDGet Then
       With vsGrid1
            .TextMatrix(kBas_Row, kBas_Col) = Format(kBas_TMai(kBas_Poket), "###")
            .Row = kBas_Row: .Col = kBas_Col
'            .CellBackColor = mNColor                                ' 2006/02/03 Delete
'----------------------------------------------------------- 2006/02/03 Insert
            If .CellBackColor <> PoketErrMsg.BackColor Then
               .CellBackColor = mNColor
            End If
'----------------------------------------------------------- 2006/02/03 Insert End
            'If kBas_Poket > 10 And kBas_Poket Mod 5 = 1 Then       ' 2006/02/03 Delete
            If kBas_Poket > 4 And kBas_Poket Mod 5 = 1 Then         ' 2006/02/03 Insert
               For gInt = kBas_Poket To kBas_Poket + 4
                    If gInt > kBas_PoketMax Then Exit For
                    kBas_TKey(gInt) = Trim(QRArea.取引先)
                    If gInt = kBas_Poket Then
                        With vsGrid2
'----------------------------------------------------------- 2006/02/03 Delete
'                            .TextMatrix(0, kBas_Col - 3) = kBas_TKey(kBas_Poket)
'                             If TorNmGet(kBas_TKey(kBas_Poket), 9) Then
'                                .TextMatrix(1, kBas_Col - 3) = CisFun.Mid2(RTrim(TRM_RName), 1, 8)
'                                .TextMatrix(2, kBas_Col - 3) = CisFun.Mid2(RTrim(TRM_RName), 9, 8)
'                             End If
'----------------------------------------------------------- 2006/02/03 Delete End
'----------------------------------------------------------- 2006/02/03 Insert
                             If kBas_Poket > 10 Then
                                If Trim(.TextMatrix(0, kBas_Col - 3)) = "" Then
                                   If TorNmGet(kBas_TKey(kBas_Poket), 9) Then
                                      .TextMatrix(0, kBas_Col - 3) = kBas_TKey(kBas_Poket)
                                      .TextMatrix(1, kBas_Col - 3) = CisFun.Mid2(RTrim(TRM_RName), 1, 8)
                                      .TextMatrix(2, kBas_Col - 3) = CisFun.Mid2(RTrim(TRM_RName), 9, 8)
                                   End If
                                End If
                             End If
                             If kBas_Poket > 5 And kBas_Poket < 11 Then
                                If Trim(B1_Name) = "" Then
                                   If TorNmGet(kBas_TKey(kBas_Poket), 9) Then
                                      B1_Name = kBas_TKey(kBas_Poket)
                                      B1_Name2(0) = CisFun.Mid2(RTrim(TRM_RName), 1, 12)
                                      B1_Name2(1) = CisFun.Mid2(RTrim(TRM_RName), 13, 12)
                                   End If
                                End If
                             End If
'----------------------------------------------------------- 2006/02/03 Insert End
                        End With
                    
                    End If
               Next gInt
            End If
       End With
    End If

KanbanProc_ED:
    kBas_ReadFlg = True
    ' 枚数表示
    lb_InCnt = CisFun.Val2(lb_InCnt) + 1
    If kBas_Poket < 9 And kBas_ErrNo <> 0 Then
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
    kBas_ColSu = 1  ' 1ﾎﾟｹｯﾄのｶﾗﾑ数
    
    GridIDGet = False
    ' かんばんポケット最大数を超えた時、ポケットオーバー
'----------------------------------------------------------- 2006/02/03 Delete
'    If kBas_Poket > kBas_PoketMax Then
'        kBas_Col = 2: kBas_Row = 4
'        GridIDGet = True
'        Exit Function
'    End If
'----------------------------------------------------------- 2006/02/03 Delete End
'----------------------------------------------------------- 2006/02/03 Insert
    If MaxCheck Then
       If kBas_Poket > kBas_PoketMax Then
          Exit Function
       End If
    End If
'----------------------------------------------------------- 2006/02/03 Insert End
    kBas_Col = kBas_Poket \ 5
    kBas_Row = kBas_Poket Mod 5
    If kBas_Row = 0 Then
        kBas_Row = 5
        kBas_Col = kBas_Col - 1
        If kBas_Col < 0 Then kBas_Col = 0
    End If
    kBas_Col = kBas_Col * kBas_ColSu + 1
    
'    If kBas_Poket > 30 Then
'        kBas_Row = kBas_Row + 6
'        kBas_Col = kBas_Col - (kBas_ColSu * 6)
'    End If
    
    GridIDGet = True
End Function
'**************************************
'*          グリッド　初期化           *
'**************************************
Private Sub GridClear()
    With vsGrid1
        .Redraw = False
        For mRId = 0 To .Rows - 1
            For mCId = 0 To .Cols - 1
                .TextMatrix(mRId, mCId) = ""
                .Row = mRId: .Col = mCId
                .CellBackColor = mMColor
            Next mCId
        Next mRId
        .Redraw = True
    End With
    With vsGrid2
        .Redraw = False
        For mRId = 0 To .Rows - 1
            For mCId = 0 To .Cols - 1
                .TextMatrix(mRId, mCId) = ""
            Next mCId
        Next mRId
        .Redraw = True
    End With
'-------------------------------------------- 2006/02/03 Insert
    Call GridColor
    B1_Name = ""
    B1_Name2(0) = ""
    B1_Name2(1) = ""
    Call PoketMaxColorSet
'-------------------------------------------- 2006/02/03 Insert End
End Sub
Private Sub GridColor()
    With vsGrid1
        For gCnt1 = 1 To .Rows - 1
            .Redraw = False
            For gCnt2 = 1 To .Cols - 1
                .Row = gCnt1: .Col = gCnt2
'                .CellBackColor = mMColor
                If .CellBackColor <> PoketErrMsg.BackColor Then
                   .CellBackColor = mMColor
                End If
            Next gCnt2
            .Redraw = True
        Next gCnt1
    End With
    
'--------------------------------------------------------- 2006/02/03 Insert
    Call PoketMaxColorSet
    Call GridErrSet
'--------------------------------------------------------- 2006/02/03 Insert End
End Sub
'----------------------------------------------------------- 2006/02/03 Insert
'+---------------------------------------+
'+      ポケット最大数を超えた時色セット
'+---------------------------------------+
Private Sub PoketMaxColorSet()

    If kBas_PoketMax < 61 And kBas_MaxPColor <> 0 Then
       kBas_PoketSave = kBas_Poket
       With vsGrid1
            .Redraw = False
            
            For gCnt1 = kBas_PoketMax + 1 To 60
                kBas_Poket = gCnt1
                Call GridIDGet(False)
                .Row = kBas_Row: .Col = kBas_Col
                .CellBackColor = kBas_MaxPColor
            Next gCnt1
            .Redraw = True
       End With
       kBas_Poket = kBas_PoketSave
    End If
End Sub
'+---------------------------------------+
'+      エラーポケット情報設定
'+---------------------------------------+
Private Function GridErrSet() As Boolean
    GridErrSet = False
    
    kBas_PoketSave = kBas_Poket
'<< かんばん読取エラーポケット表示 >>
    For Each gObj In B1_ErrName
        gObj.Visible = False
    Next gObj
    
    For gLong = 1 To UBound(kBas_ErSetMsg)
        If Trim(kBas_ErSetMsg(gLong)) <> "" Then
           If gLong < 11 Then
              B1_ErrName(gLong) = Trim(kBas_ErSetMsg(gLong))
              B1_ErrName(gLong).Visible = True
           Else
              kBas_Poket = gLong
              If GridIDGet Then
                 With vsGrid1
                     .Row = kBas_Row: .Col = kBas_Col
                     .CellBackColor = PoketErrMsg.BackColor
                     .CellAlignment = flexAlignCenterCenter
                 End With
              End If
           End If
        End If
    Next gLong
    
    kBas_Poket = kBas_PoketSave
    
    GridErrSet = True
End Function
'----------------------------------------------------------- 2006/02/03 Insert End

