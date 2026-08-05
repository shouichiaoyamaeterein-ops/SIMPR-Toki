VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{490AF35D-478A-11D4-AEC2-00000EA57F9E}#4.0#0"; "CISMSG~2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CKH1040 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "発注かんばん入力 [指定] 仕入先単位"
   ClientHeight    =   10650
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   14985
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
   ScaleHeight     =   10650
   ScaleWidth      =   14985
   WindowState     =   2  '最大化
   Begin VB.PictureBox PNL_Start 
      BackColor       =   &H00000000&
      Height          =   825
      Left            =   4605
      ScaleHeight     =   765
      ScaleWidth      =   7125
      TabIndex        =   9
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
      TabIndex        =   6
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
               TabIndex        =   7
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
               TabIndex        =   8
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
               BackColorBkg    =   12632256
               FocusRect       =   0
               HighLight       =   0
               ScrollBars      =   0
               RowHeightMin    =   280
               GridLines       =   9
               GridLinesFixed  =   9
               SelectionMode   =   1
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
      TabIndex        =   5
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
      Caption         =   "【 発注かんばん入力　[指定] 仕入先単位 】"
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
      Top             =   510
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
      Top             =   1500
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
   Begin Cis3D_v60.CIS3D H1_Area1 
      Height          =   615
      Left            =   360
      Top             =   1545
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   1085
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
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   510
         Left            =   60
         Top             =   60
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   900
         ForeColor       =   16711680
         Caption         =   "発注区分"
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
         Begin VB.OptionButton OpH1_HKbn 
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
            Index           =   2
            Left            =   3300
            TabIndex        =   4
            Top             =   105
            Width           =   1335
         End
         Begin VB.OptionButton OpH1_HKbn 
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
            Index           =   3
            Left            =   2235
            TabIndex        =   3
            Top             =   105
            Width           =   990
         End
         Begin VB.OptionButton OpH1_HKbn 
            BackColor       =   &H00C0C0C0&
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
            TabIndex        =   2
            Top             =   105
            Value           =   -1  'True
            Width           =   990
         End
      End
   End
   Begin Cis3D_v60.CIS3D H1_Area2 
      Height          =   870
      Left            =   345
      Top             =   2235
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   1535
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
      Begin Cis3D_v60.CIS3D Back_Ymd 
         Height          =   765
         Left            =   60
         Top             =   60
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "発 注 日"
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_Ymd 
            Height          =   330
            Left            =   75
            TabIndex        =   0
            Top             =   315
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   582
            cButton         =   -1  'True
            cCalender       =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
      End
      Begin Cis3D_v60.CIS3D CIS3D21 
         Height          =   765
         Left            =   2235
         Top             =   60
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1349
         ForeColor       =   16711680
         Caption         =   "便"
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
         Begin CisText_V60.CisText iH1_Bin 
            Height          =   360
            Left            =   150
            TabIndex        =   1
            Top             =   300
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
   End
End
Attribute VB_Name = "CKH1040"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   発注かんばん入力[指定]仕入先単位
'**       フォームID    :   CKH1040
'**       処理概要      :
'**
'**       作  成  日    :   2004/01/26  By CIS SIMPR-A
'**       変  更  日    :   2004/09/17  BY CIS 部品材用混在可能
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
        
    iH1_Ymd = Format(Date, "YYYYMMDD")
    
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
    ProcHB = "H1"
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
    
' 部品・材料
    kBas_Proc = "H"
    gStr = UCase(Command)
    If InStr(1, gStr, "/B") <> 0 Or InStr(1, gStr, "/部品") <> 0 Then
       kBas_Proc = ""
       If InStr(1, gStr, "/C") = 0 Then
          PGName = Replace(PGName, "発注", "部品")
       End If
    End If
    If InStr(1, gStr, "/Z") <> 0 Or InStr(1, gStr, "/材料") <> 0 Then
       kBas_Proc = "Z"
       PGName.BackColor = &H80C0FF
       If InStr(1, gStr, "/C") = 0 Then
          PGName = Replace(PGName, "発注", "材料")
       End If
       Head_Tanto.BackColor = &H80C0FF
       Pgid.BackColor = &H80C0FF
       Pgid.cBoderColor1 = &H80C0FF
       Pgid.cBoderColor2 = &H80C0FF
    End If
    
    If Not cKanbanInit Then
        Call PB_END_Click
    End If
'初期処理・初期表示
    If Not Init_Set Then
        Call PB_END_Click
    End If
    
    If kBas_ExecMode Then
       K_Sykbnm.Visible = True
    Else
       K_Sykbnm.Visible = False
    End If
'
    Call OpH1_HKbn_Click(1)
    
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
           Case vbKeyF7:       If ProcHB = "H1" And Trim(iH1_Ymd.Tag) <> "" Then iH1_Ymd.ShowCalender  '【終了】
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
'+   後方処理(部品)
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
           CisDB.PassLink
        ' 発注処理
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
'+   後方処理(材料)
'+---------------------+
    If kBas_Check.ErrCntZ <> 0 Or kBas_Check.OutCntZ <> 0 Then
       CKHZ0900.Show vbModal
       Unload CKHZ0900
       Set CKHZ0900 = Nothing
       DoEvents
        ' 発注処理対象データ確認
       gSL_Select = "Select Count(*) 件数 From 材料発注かんばん予定テーブル "
       gSL_Select = gSL_Select & " Where 処理端末 = Host_Name()"
       gSL_Select = gSL_Select & "   And 処理区分 = 1"
       Call IRNRead(gSL_Select, 1)
       If IRN.件数 <> 0 Then
           CisDB.PassLink
        ' 発注処理
           On Error Resume Next
           With CisFun
                .ShellApl = gIniExe & "PXHZ0110.exe"
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
   Call ReturnPress
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
'+----------------------------+
'+  OpH1_HKbn  発注区分        +
'+----------------------------+
Private Sub OpH1_HKbn_Click(Index As Integer)
    OpH1_HKbn(1).BackColor = gPLostSel
    OpH1_HKbn(2).BackColor = gPLostSel
    OpH1_HKbn(3).BackColor = gPLostSel
    OpH1_HKbn(Index).BackColor = &HC0FFC0
    OpH1_HKbn(1).Tag = Index
    If Index = 1 Then
        Back_Ymd = "発 注 日"
    Else
        Back_Ymd = "納 入 日"
    End If
End Sub
'+-----------------------------+
'+  iH1_Ymd
'+-----------------------------+
Private Sub iH1_Ymd_GotFocus()
    iH1_Ymd.Tag = ActiveControl.Name
End Sub
Private Sub iH1_Ymd_LostFocus()
    iH1_Ymd.Tag = ""
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
    Dim RDMode      As Boolean
    Dim RDColor     As Long
    H1Mode = False
    H1Color = gPLostSel
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
       Case "RD"
          RDMode = True
          RDColor = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    H1_Area2.Enabled = H1Mode
    H1_Area2.BackColor = H1Color
    PNL_PF.Visible = H1Mode
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
   Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
       If Head1Chk Then
          Call KanbanRead
       End If
       GoTo ReturnPress_Ed
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'*------------------------------------------------------------------------'
'*          Head1 チェック
'*------------------------------------------------------------------------'
Private Function Head1Chk() As Boolean
    Head1Chk = False
'日付チェック
    If RTrim(iH1_Ymd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "     　【" & RTrim(Back_Ymd) & " 必須！ 】     "
            .MB_MSG(3) = "     " & RTrim(Back_Ymd) & " を入力して下さい。      "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_Ymd.SetFocus
        Exit Function
    End If
    If Not iH1_Ymd.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "     　【" & RTrim(Back_Ymd) & " 異常！ 】     "
            .MB_MSG(3) = "     " & RTrim(Back_Ymd) & " を再入力して下さい。      "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_Ymd.SetFocus
        Exit Function
    End If
    If RTrim(iH1_Ymd) < Format(Now(), "yyyymmdd") Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "     " & RTrim(Back_Ymd) & " が過去の日付です。     "
            .MB_MSG(3) = "     " & RTrim(Back_Ymd) & " を再入力して下さい。      "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_Ymd.SetFocus
        Exit Function
    End If
'便チェック
    If CisFun.Val2(iH1_Bin) = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "     　【　便 必須！ 】     "
            .MB_MSG(3) = "     便を入力して下さい。      "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_Bin.SetFocus
        Exit Function
    End If
    
    Head1Chk = True
End Function
'+---------------------------------------+
'+      初期設定・初期表示
'+---------------------------------------+
Private Function Init_Set() As Boolean
    Init_Set = False
'<< 発注区分 >>
    gSL_Select = " SELECT * FROM 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = '発注区分'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   And 値 < '4'"
    gSL_Select = gSL_Select & " ORDER BY 値"
    If SYMRead(gSL_Select, 0) Then
        Do Until Not SYM_RDSTS
            OpH1_HKbn(CisFun.Val2(SYM.値)).Caption = RTrim(SYM.値名称)
            Call SYMReadNext
        Loop
    End If
    Call SYMClose
'<< かんばん読取エラーポケット表示 >>
    gSL_Select = " SELECT * FROM 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = 'かんばんエラーポケット'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   And 数字1 = 0"
    gSL_Select = gSL_Select & " ORDER BY 値"
    If SYMRead(gSL_Select, 0) Then
        gInt = 0
        Do Until Not SYM_RDSTS
            gInt = gInt + 1
            B1_ErrName(gInt) = RTrim(SYM.値名称)
            Call SYMReadNext
        Loop
    End If
    Call SYMClose
    
    Init_Set = True
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
'       Call TX597GClose
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
    
    ' 部品指定時、部品かんばん以外はエラー
    If kBas_Proc = "B" And QRArea.品目 <> 1 Then
       kBas_Poket = kBas_PokErr
       kBas_ErrNo = 99
       GoTo KanbanProc_Set:
    End If
    ' 材料指定時、材料かんばん以外はエラー
    If kBas_Proc = "Z" And QRArea.品目 <> 2 Then
       kBas_Poket = kBas_PokErr
       kBas_ErrNo = 99
       GoTo KanbanProc_Set:
    End If
    ' 部品かんばん・材料かんばん以外はエラー
    If QRArea.品目 <> 1 And QRArea.品目 <> 2 Then
       kBas_Poket = kBas_PokErr
       kBas_ErrNo = 99
       GoTo KanbanProc_Set:
    End If
    
    mPoketFlg = False
    
    ' 部品かんばん読取内容チェック
    If QRArea.品目 = 1 Then
        With kBas_Check
            .ProcKB = "2"
            .HCKB = CisFun.Val2(OpH1_HKbn(1).Tag)
            .HYMD = iH1_Ymd
            .HBIN = Format(CisFun.Val2(iH1_Bin), "00")
        End With
        If Not cKanbanChkH Then
           KanbanProc = False
           kBas_ErrNo = 99
           GoTo KanbanProc_Set
        End If
    End If
    
    ' 材料かんばん読取内容チェック
    If QRArea.品目 = 2 Then
        With kBas_Check
            .ProcKB = "2"
            .HCKB = CisFun.Val2(OpH1_HKbn(1).Tag)
            .HYMD = iH1_Ymd
            .HBIN = Format(CisFun.Val2(iH1_Bin), "00")
        End With
        If Not cKanbanChkHZ Then
           KanbanProc = False
           kBas_ErrNo = 99
           GoTo KanbanProc_Set
        End If
    End If
    If kBas_Poket <> 0 Then GoTo KanbanProc_Set
    
    For kBas_PCnt = 11 To kBas_PoketMax
        mPSetFlg = False
        If Trim(kBas_TKey(kBas_PCnt)) = Trim(mKey) Or _
           Trim(kBas_TKey(kBas_PCnt)) = "" Then
           mPSetFlg = True
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
       If kBas_TMai(9) < kBas_TMaiMax(5) Or _
          kBas_TMaiMax(5) = 0 Then
          kBas_Poket = 9
       Else
          kBas_Poket = 10
       End If
    End If
    '
KanbanProc_Set:
' ｸﾞﾘｯﾄﾞへ内容セット & 枚数カウント
    kBas_TMai(kBas_Poket) = kBas_TMai(kBas_Poket) + 1
    
    Call GridColor
    
    If GridIDGet Then
       With vsGrid1
            .TextMatrix(kBas_Row, kBas_Col) = Format(kBas_TMai(kBas_Poket), "###")
            .Row = kBas_Row: .Col = kBas_Col
            .CellBackColor = mNColor
            If kBas_Poket > 10 And kBas_Poket Mod 5 = 1 Then
               For gInt = kBas_Poket To kBas_Poket + 4
                    If gInt > kBas_PoketMax Then Exit For
                    kBas_TKey(gInt) = Trim(QRArea.取引先)
                    If gInt = kBas_Poket Then
                        With vsGrid2
                            .TextMatrix(0, kBas_Col - 3) = kBas_TKey(kBas_Poket)
                             If TorNmGet(kBas_TKey(kBas_Poket), 9) Then
                                .TextMatrix(1, kBas_Col - 3) = CisFun.Mid2(RTrim(TRM_RName), 1, 8)
                                .TextMatrix(2, kBas_Col - 3) = CisFun.Mid2(RTrim(TRM_RName), 9, 8)
                             End If
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
    If kBas_ErrNo <> 0 And kBas_Poket < 9 Then
        lb_ErrCnt = Format(CisFun.Val2(lb_ErrCnt) + 1, "#,###")
    Else
        lb_OKCnt = Format(CisFun.Val2(lb_OKCnt) + 1, "#,###")
    End If

End Function
'+------------------------------------+
'+      グリッド位置(Row,Col)を求める
'+------------------------------------+
Private Function GridIDGet() As Boolean
    kBas_ColSu = 1  ' 1ﾎﾟｹｯﾄのｶﾗﾑ数
    
    GridIDGet = False
    ' かんばんポケット最大数を超えた時、ポケットオーバー
    If kBas_Poket > kBas_PoketMax Then
        kBas_Col = 2: kBas_Row = 4
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
End Sub
Private Sub GridColor()
    With vsGrid1
        For gCnt1 = 1 To .Rows - 1
            .Redraw = False
            For gCnt2 = 1 To .Cols - 1
                .Row = gCnt1: .Col = gCnt2
                .CellBackColor = mMColor
            Next gCnt2
            .Redraw = True
        Next gCnt1
    End With
End Sub


