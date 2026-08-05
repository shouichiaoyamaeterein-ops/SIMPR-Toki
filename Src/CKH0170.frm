VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CisYMDwB_6.0.ocx"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CKH0170 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "かんばん管理入力(増減・平準化)"
   ClientHeight    =   11235
   ClientLeft      =   135
   ClientTop       =   405
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
   ForeColor       =   &H00C0C000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
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
      Left            =   0
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   0
      Width           =   585
   End
   Begin Cis3D_v60.CIS3D B1_Area3 
      Height          =   4440
      Left            =   8655
      Top             =   6225
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   7832
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
      Begin Cis3D_v60.CIS3D CIS3D16 
         Height          =   1860
         Left            =   60
         Top             =   2520
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3281
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
         Begin Cis3D_v60.CIS3D CIS3D27 
            Height          =   420
            Index           =   0
            Left            =   30
            Top             =   30
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   741
            Caption         =   ""
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D12 
            Height          =   420
            Left            =   1755
            Top             =   30
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   741
            Caption         =   "発行"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B_Color 
            Height          =   420
            Index           =   1
            Left            =   2625
            Top             =   30
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   741
            BackColor       =   16761024
            Caption         =   "有効"
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
            cAlingnment     =   7
            cBoderStyle     =   1
            cPositionX      =   15
         End
         Begin Cis3D_v60.CIS3D B_Color 
            Height          =   420
            Index           =   2
            Left            =   3495
            Top             =   30
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   741
            BackColor       =   16761087
            Caption         =   "無効"
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
            cAlingnment     =   7
            cBoderStyle     =   1
            cPositionX      =   15
         End
         Begin Cis3D_v60.CIS3D B_Color 
            Height          =   420
            Index           =   3
            Left            =   4365
            Top             =   30
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   741
            BackColor       =   16711935
            Caption         =   "滞留"
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
            cAlingnment     =   7
            cBoderStyle     =   1
            cPositionX      =   15
         End
         Begin Cis3D_v60.CIS3D B_Color 
            Height          =   420
            Index           =   0
            Left            =   5235
            Top             =   30
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   741
            BackColor       =   8421631
            Caption         =   "未流動"
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
            cAlingnment     =   7
            cBoderStyle     =   1
            cPositionX      =   15
         End
         Begin Cis3D_v60.CIS3D B1lb_HMaisu 
            Height          =   540
            Index           =   0
            Left            =   1755
            Top             =   450
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   953
            BackColor       =   16777215
            Caption         =   "###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_YSu 
            Height          =   540
            Index           =   0
            Left            =   2625
            Top             =   450
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   953
            BackColor       =   16777215
            Caption         =   "###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_MSu 
            Height          =   540
            Index           =   0
            Left            =   3495
            Top             =   450
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   953
            BackColor       =   16777215
            Caption         =   "###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_TSu 
            Height          =   540
            Index           =   0
            Left            =   4365
            Top             =   450
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   953
            BackColor       =   16777215
            Caption         =   "###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_MISu 
            Height          =   540
            Index           =   0
            Left            =   5235
            Top             =   450
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   953
            BackColor       =   16777215
            Caption         =   "###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_YSu 
            Height          =   405
            Index           =   1
            Left            =   2625
            Top             =   990
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   714
            BackColor       =   12648447
            Caption         =   "###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_MSu 
            Height          =   405
            Index           =   1
            Left            =   3495
            Top             =   990
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   714
            BackColor       =   12648447
            Caption         =   "###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_TSu 
            Height          =   405
            Index           =   1
            Left            =   4365
            Top             =   990
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   714
            BackColor       =   12648447
            Caption         =   "###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_MISu 
            Height          =   405
            Index           =   1
            Left            =   5235
            Top             =   990
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   714
            BackColor       =   12648447
            Caption         =   "###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_YSu 
            Height          =   405
            Index           =   2
            Left            =   2625
            Top             =   1395
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   714
            BackColor       =   12648384
            Caption         =   "###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_MSu 
            Height          =   405
            Index           =   2
            Left            =   3495
            Top             =   1395
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   714
            BackColor       =   12648384
            Caption         =   "###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_TSu 
            Height          =   405
            Index           =   2
            Left            =   4365
            Top             =   1395
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   714
            BackColor       =   12648384
            Caption         =   "###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_MISu 
            Height          =   405
            Index           =   2
            Left            =   5235
            Top             =   1395
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   714
            BackColor       =   12648384
            Caption         =   "###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_HMaisu 
            Height          =   405
            Index           =   1
            Left            =   1755
            Top             =   990
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   714
            BackColor       =   12648447
            Caption         =   "###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_HMaisu 
            Height          =   405
            Index           =   2
            Left            =   1755
            Top             =   1395
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   714
            BackColor       =   12648384
            Caption         =   "###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D CIS3D27 
            Height          =   540
            Index           =   1
            Left            =   30
            Top             =   450
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   953
            BackColor       =   16777215
            Caption         =   "合計枚数"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   14.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D27 
            Height          =   405
            Index           =   2
            Left            =   30
            Top             =   990
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   714
            BackColor       =   12648447
            Caption         =   "収容数①"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   14.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D27 
            Height          =   405
            Index           =   3
            Left            =   30
            Top             =   1395
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   714
            BackColor       =   12648384
            Caption         =   "収容数②"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   14.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D10 
         Height          =   480
         Index           =   0
         Left            =   45
         Top             =   105
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   847
         BackColor       =   4210752
         ForeColor       =   65535
         Caption         =   "マ ス タ 情 報"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
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
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   1545
         Index           =   0
         Left            =   60
         Top             =   585
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   2725
         ForeColor       =   0
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
         cAlingnment     =   6
         Begin Cis3D_v60.CIS3D B1lb_HKeisu 
            Height          =   345
            Left            =   810
            Top             =   1170
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   609
            ForeColor       =   0
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
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B1lb_StartDay 
            Height          =   375
            Left            =   810
            Top             =   30
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   661
            BackColor       =   16761024
            ForeColor       =   0
            Caption         =   "-1"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   345
            Index           =   0
            Left            =   1395
            Top             =   1170
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   609
            Caption         =   "変 更 日 ･ 便"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   360
            Left            =   1395
            Top             =   30
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   635
            Caption         =   "ｻｲｸﾙ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B_Syuyo 
            Height          =   330
            Index           =   1
            Left            =   3870
            Top             =   45
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   582
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
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   360
            Left            =   2295
            Top             =   30
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   635
            Caption         =   "置  場"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B1lb_Cycle 
            Height          =   780
            Left            =   1395
            Top             =   390
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   1376
            BackColor       =   14737632
            Caption         =   "XXXXXXX"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B1lb_Okiba 
            Height          =   780
            Left            =   2295
            Top             =   390
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   1376
            BackColor       =   14737632
            Caption         =   "XXXXXXXXX1"
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
            cPositionX      =   60
         End
         Begin Cis3D_v60.CIS3D B1lb_Syuyo 
            Height          =   390
            Index           =   2
            Left            =   3870
            Top             =   780
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   688
            BackColor       =   12648384
            Caption         =   "#,###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_SYmd 
            Height          =   345
            Left            =   3570
            Top             =   1170
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   609
            BackColor       =   14737632
            Caption         =   "YYYY/MM/DD 99"
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
            cAlingnment     =   7
            cBoderStyle     =   1
            cPositionX      =   60
         End
         Begin Cis3D_v60.CIS3D B1lb_Syuyo 
            Height          =   390
            Index           =   1
            Left            =   3870
            Top             =   390
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   688
            BackColor       =   12648447
            Caption         =   "###,###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B_Syuyo 
            Height          =   330
            Index           =   0
            Left            =   4830
            Top             =   45
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   582
            Caption         =   "収容器"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B1lb_Syuki 
            Height          =   390
            Index           =   1
            Left            =   4830
            Top             =   390
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   688
            BackColor       =   12648447
            Caption         =   "XXXXXXXXX1"
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
            cPositionX      =   60
         End
         Begin Cis3D_v60.CIS3D B1lb_Syuki 
            Height          =   390
            Index           =   2
            Left            =   4830
            Top             =   780
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   688
            BackColor       =   12648384
            Caption         =   "XXXXXXXXX1"
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
            cPositionX      =   60
         End
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   390
            Index           =   3
            Left            =   3570
            Top             =   390
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   688
            BackColor       =   12648447
            ForeColor       =   0
            Caption         =   "①"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   390
            Index           =   4
            Left            =   3570
            Top             =   780
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   688
            BackColor       =   12648384
            ForeColor       =   0
            Caption         =   "②"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B_Syuyo 
            Height          =   330
            Index           =   4
            Left            =   3570
            Top             =   45
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   582
            Caption         =   ""
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   345
            Index           =   19
            Left            =   30
            Top             =   810
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   609
            ForeColor       =   0
            Caption         =   "平準化"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   30
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   375
            Index           =   0
            Left            =   30
            Top             =   420
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   661
            ForeColor       =   0
            Caption         =   "か係数"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   30
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   375
            Index           =   1
            Left            =   30
            Top             =   30
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   661
            ForeColor       =   0
            Caption         =   "開始日"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   30
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   345
            Index           =   2
            Left            =   30
            Top             =   1170
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   609
            ForeColor       =   0
            Caption         =   "平係数"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   30
         End
         Begin Cis3D_v60.CIS3D B1lb_KKeisu 
            Height          =   375
            Left            =   810
            Top             =   420
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   661
            ForeColor       =   0
            Caption         =   "9.99"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B1lb_HKbn 
            Height          =   345
            Left            =   810
            Top             =   810
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   609
            ForeColor       =   0
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
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D10 
         Height          =   360
         Index           =   1
         Left            =   30
         Top             =   2130
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   635
         BackColor       =   4210752
         ForeColor       =   65535
         Caption         =   "か ん ば ん 管 理 情 報"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
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
   End
   Begin Cis3D_v60.CIS3D B1_Area2 
      Height          =   4395
      Left            =   360
      Top             =   6255
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   7752
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
      Begin Cis3D_v60.CIS3D CIS3D25 
         Height          =   435
         Index           =   0
         Left            =   105
         Top             =   105
         Width           =   7950
         _ExtentX        =   14023
         _ExtentY        =   767
         BackColor       =   4210752
         ForeColor       =   12632256
         Caption         =   "当 月 内 示 情 報"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
         Begin Cis3D_v60.CIS3D CIS3D25 
            Height          =   315
            Index           =   1
            Left            =   5580
            Top             =   135
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            BackColor       =   4210752
            ForeColor       =   12632256
            Caption         =   "()内は前回情報"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   4210752
            cBoderColor2    =   4210752
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   1335
         Left            =   105
         Top             =   540
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   2355
         BackColor       =   16761024
         ForeColor       =   4210752
         Caption         =   "年 月"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   90
         Begin Cis3D_v60.CIS3D B1lb_NYM 
            Height          =   810
            Left            =   75
            Top             =   435
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   1429
            BackColor       =   14737632
            ForeColor       =   192
            Caption         =   "9999年99月"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   2025
         Left            =   120
         Top             =   2280
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   3572
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
         Begin vsFlexLib.vsFlexArray vsGrid2 
            Height          =   1905
            Left            =   30
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   30
            Width           =   7800
            _Version        =   196608
            _ExtentX        =   13758
            _ExtentY        =   3360
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
            FormatString    =   "№ |内示年月|内示数   |日当枚数 |回転枚数 |前月差|平準化日 |平準化便 |"
            Rows            =   6
            Cols            =   11
            BackColor       =   16777152
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   0
            RowHeightMin    =   300
            GridLines       =   3
            SelectionMode   =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   1335
         Left            =   2280
         Top             =   540
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   2355
         BackColor       =   16761024
         ForeColor       =   4210752
         Caption         =   "内 示 数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   90
         Begin Cis3D_v60.CIS3D B1lb_Naisu 
            Height          =   810
            Left            =   75
            Top             =   450
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   1429
            BackColor       =   14737632
            ForeColor       =   192
            Caption         =   "999,999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   3
            cBoderStyle     =   1
            cPositionX      =   -30
            cPositionY      =   90
            Begin Cis3D_v60.CIS3D B1lb_ZenNaisu 
               Height          =   315
               Left            =   195
               Top             =   465
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   556
               BackColor       =   14737632
               ForeColor       =   12632064
               Caption         =   "(999,999,999)"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   14737632
               cBoderColor2    =   14737632
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   3
               cBoderStyle     =   2
            End
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D17 
         Height          =   1335
         Left            =   4560
         Top             =   540
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   2355
         BackColor       =   16761024
         ForeColor       =   4210752
         Caption         =   "日当枚数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   90
         Begin Cis3D_v60.CIS3D B1lb_Daysu 
            Height          =   810
            Left            =   75
            Top             =   450
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   1429
            BackColor       =   14737632
            ForeColor       =   192
            Caption         =   "999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   90
            Begin Cis3D_v60.CIS3D B1lb_ZenDaysu 
               Height          =   315
               Left            =   165
               Top             =   465
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   556
               BackColor       =   14737632
               ForeColor       =   12632064
               Caption         =   "(999)"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   14737632
               cBoderColor2    =   14737632
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cBoderStyle     =   2
            End
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D21 
         Height          =   1335
         Left            =   5850
         Top             =   540
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   2355
         BackColor       =   16761024
         ForeColor       =   4210752
         Caption         =   "回転枚数 / 差"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   90
         Begin Cis3D_v60.CIS3D B1lb_KaiMai 
            Height          =   810
            Left            =   60
            Top             =   450
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1429
            BackColor       =   14737632
            ForeColor       =   192
            Caption         =   "999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   90
            Begin Cis3D_v60.CIS3D B1lb_ZenKaiMai 
               Height          =   315
               Left            =   30
               Top             =   465
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   556
               BackColor       =   14737632
               ForeColor       =   12632064
               Caption         =   "(999)"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   14737632
               cBoderColor2    =   14737632
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cBoderStyle     =   2
            End
         End
         Begin Cis3D_v60.CIS3D B1lb_KaiMaiSa 
            Height          =   810
            Left            =   900
            Top             =   450
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   1429
            BackColor       =   14737632
            ForeColor       =   192
            Caption         =   "-999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   20.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
            Begin VB.Shape Shape1 
               BorderColor     =   &H000000C0&
               BorderWidth     =   3
               Height          =   690
               Left            =   60
               Shape           =   4  '丸みのある長方形
               Top             =   60
               Width           =   1125
            End
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D26 
         Height          =   435
         Left            =   120
         Top             =   1890
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   767
         BackColor       =   4210752
         ForeColor       =   14737632
         Caption         =   "内 示 情 報 ( 履歴 )"
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
         cAlingnment     =   7
      End
   End
   Begin Cis3D_v60.CIS3D B1_Area9 
      Height          =   3660
      Left            =   360
      Top             =   2550
      Width           =   14610
      _ExtentX        =   25770
      _ExtentY        =   6456
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
      Begin VB.PictureBox B1_Area4 
         Height          =   1980
         Left            =   30
         ScaleHeight     =   1920
         ScaleWidth      =   14475
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1635
         Width           =   14535
         Begin vsFlexLib.vsFlexArray vsGrid1 
            Height          =   1890
            Left            =   45
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   15
            Width           =   14400
            _Version        =   196608
            _ExtentX        =   25400
            _ExtentY        =   3334
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
            FormatString    =   "№|設定 |開 始 日   |終 了 日   |増減 |設定枚数 |便当枚数 |累積枚数 |完 了 日  |平準化|上限枚数|更新日時|"
            Rows            =   6
            Cols            =   13
            BackColor       =   12648447
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   300
            GridLines       =   3
            SelectionMode   =   1
            AllowBigSelection=   0   'False
            BackColorAlternate=   12648447
         End
      End
      Begin VB.PictureBox B1_Area1 
         Height          =   1575
         Left            =   30
         ScaleHeight     =   1515
         ScaleWidth      =   14475
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   30
         Width           =   14535
         Begin Cis3D_v60.CIS3D CIS3D28 
            Height          =   1425
            Left            =   45
            Top             =   45
            Width           =   14355
            _ExtentX        =   25321
            _ExtentY        =   2514
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
               Height          =   1020
               Index           =   0
               Left            =   13320
               Top             =   390
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   1799
               ForeColor       =   0
               Caption         =   "増減結果"
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
               cPositionY      =   120
               Begin Cis3D_v60.CIS3D B1lb_AfterHeiSu 
                  Height          =   360
                  Left            =   150
                  Top             =   480
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "#,###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -60
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   1020
               Index           =   2
               Left            =   3900
               Top             =   390
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   1799
               ForeColor       =   16711680
               Caption         =   "区分"
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
               cPositionX      =   30
               cPositionY      =   30
               Begin VB.OptionButton B1Op_ZGKbn 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "設定なし"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   0
                  Left            =   60
                  TabIndex        =   9
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   1350
               End
               Begin VB.OptionButton B1Op_ZGKbn 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "減"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   60
                  TabIndex        =   10
                  Top             =   690
                  Width           =   570
               End
               Begin VB.OptionButton B1Op_ZGKbn 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "増"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   2
                  Left            =   810
                  TabIndex        =   11
                  Top             =   690
                  Width           =   570
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   690
               Index           =   3
               Left            =   1080
               Top             =   30
               Width           =   2820
               _ExtentX        =   4974
               _ExtentY        =   1217
               ForeColor       =   16711680
               Caption         =   "開始日"
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
               cPositionX      =   60
               Begin CisYMD_With_Btn.CisYMDwB iB1_SKYmd 
                  Height          =   360
                  Left            =   840
                  TabIndex        =   7
                  Top             =   195
                  Width           =   1860
                  _ExtentX        =   3069
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
                  Object.Width           =   1860
                  cSize           =   1
                  cChkResult      =   0   'False
                  cTypeIntval     =   0
                  BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   1020
               Index           =   5
               Left            =   5430
               Top             =   390
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   1799
               ForeColor       =   16711680
               Caption         =   "設定枚数"
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
               cPositionX      =   30
               cPositionY      =   120
               Begin CisText_V60.CisText iB1_SetSu 
                  Height          =   360
                  Left            =   300
                  TabIndex        =   12
                  Top             =   540
                  Width           =   510
                  _ExtentX        =   900
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
                  Text            =   "###"
                  MaxLength       =   3
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   1020
               Index           =   7
               Left            =   7530
               Top             =   390
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   1799
               ForeColor       =   0
               Caption         =   "累積枚数"
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
               cPositionX      =   30
               cPositionY      =   120
               Begin Cis3D_v60.CIS3D B1lb_RuiSu 
                  Height          =   360
                  Left            =   270
                  Top             =   510
                  Width           =   510
                  _ExtentX        =   900
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -60
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   1380
               Index           =   1
               Left            =   30
               Top             =   30
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   2434
               ForeColor       =   16711680
               Caption         =   "設定"
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
               cPositionX      =   30
               cPositionY      =   120
               Begin VB.CheckBox B1Chk_SKbn 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "有効"
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
                  Left            =   150
                  TabIndex        =   6
                  Top             =   720
                  Width           =   810
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   1020
               Index           =   6
               Left            =   6480
               Top             =   390
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   1799
               ForeColor       =   16711680
               Caption         =   "便当枚数"
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
               cPositionX      =   30
               cPositionY      =   120
               Begin CisText_V60.CisText iB1_BinSu 
                  Height          =   360
                  Left            =   270
                  TabIndex        =   13
                  Top             =   525
                  Width           =   510
                  _ExtentX        =   900
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
                  Text            =   "###"
                  MaxLength       =   3
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   1020
               Index           =   8
               Left            =   8580
               Top             =   390
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   1799
               ForeColor       =   0
               Caption         =   "完 了 日"
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
               cPositionY      =   120
               Begin Cis3D_v60.CIS3D B1lb_KanYmd 
                  Height          =   360
                  Left            =   90
                  Top             =   510
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "9999/99/99"
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
                  cAlingnment     =   7
                  cBoderStyle     =   1
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   690
               Index           =   4
               Left            =   1080
               Top             =   720
               Width           =   2820
               _ExtentX        =   4974
               _ExtentY        =   1217
               ForeColor       =   16711680
               Caption         =   "終了日"
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
               cPositionX      =   60
               Begin CisYMD_With_Btn.CisYMDwB iB1_EKYmd 
                  Height          =   360
                  Left            =   840
                  TabIndex        =   8
                  Top             =   195
                  Width           =   1860
                  _ExtentX        =   3069
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
                  Object.Width           =   1860
                  cSize           =   1
                  cChkResult      =   0   'False
                  cTypeIntval     =   0
                  BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   1020
               Index           =   7
               Left            =   12255
               Top             =   390
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   1799
               ForeColor       =   16711680
               Caption         =   "上限枚数"
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
               cPositionY      =   120
               Begin CisText_V60.CisText iB1_MonSu 
                  Height          =   360
                  Left            =   210
                  TabIndex        =   18
                  Top             =   495
                  Width           =   735
                  _ExtentX        =   1296
                  _ExtentY        =   635
                  cFormat         =   "#,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#,###"
                  MaxLength       =   5
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "####"
                  cILength        =   4
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   1020
               Index           =   9
               Left            =   10125
               Top             =   390
               Width           =   2130
               _ExtentX        =   3757
               _ExtentY        =   1799
               ForeColor       =   16711680
               Caption         =   "区分"
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
               cPositionX      =   30
               cPositionY      =   30
               Begin VB.OptionButton B1Op_HJKbn 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "設定なし"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   0
                  Left            =   90
                  TabIndex        =   14
                  Top             =   390
                  Value           =   -1  'True
                  Width           =   1350
               End
               Begin VB.OptionButton B1Op_HJKbn 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "月"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   3
                  Left            =   1500
                  TabIndex        =   17
                  Top             =   690
                  Width           =   570
               End
               Begin VB.OptionButton B1Op_HJKbn 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "日"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   2
                  Left            =   810
                  TabIndex        =   16
                  Top             =   690
                  Width           =   570
               End
               Begin VB.OptionButton B1Op_HJKbn 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "便"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   90
                  TabIndex        =   15
                  Top             =   690
                  Width           =   570
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D48 
               Height          =   360
               Index           =   1
               Left            =   3900
               Top             =   30
               Width           =   6230
               _ExtentX        =   10980
               _ExtentY        =   635
               BackColor       =   0
               ForeColor       =   65535
               Caption         =   "増減管理"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D CIS3D48 
               Height          =   360
               Index           =   2
               Left            =   10110
               Top             =   30
               Width           =   4230
               _ExtentX        =   7461
               _ExtentY        =   635
               BackColor       =   65535
               ForeColor       =   0
               Caption         =   "平準化管理"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   1
               cAlingnment     =   7
            End
         End
      End
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
      Height          =   765
      Left            =   7290
      ScaleHeight     =   705
      ScaleWidth      =   7620
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   480
      Width           =   7680
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   5061
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         Left            =   5897
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         Left            =   6735
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         Left            =   4225
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         Left            =   881
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   585
         Left            =   1717
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   7
         BtnCaption      =   "新規"
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
         Left            =   2553
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   3389
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
      Height          =   1110
      Left            =   360
      ScaleHeight     =   1050
      ScaleWidth      =   10710
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1350
      Width           =   10770
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   960
         Index           =   0
         Left            =   60
         Top             =   60
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1693
         ForeColor       =   192
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   60
         Begin Cis3D_v60.CIS3D H1lb_Seban 
            Height          =   345
            Left            =   90
            Top             =   450
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   609
            BackColor       =   16777152
            Caption         =   "XXXXXXXX"
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
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin CisText_V60.CisText iH1_Seban 
            Height          =   360
            Left            =   90
            TabIndex        =   0
            Top             =   450
            Width           =   1110
            _ExtentX        =   1958
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
            Text            =   "XXXXXXXX"
            MaxLength       =   8
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D BLB_HINBN 
         Height          =   960
         Left            =   5715
         Top             =   60
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1693
         ForeColor       =   16711680
         Caption         =   "品  番"
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
         cPositionY      =   60
         Begin Cis3D_v60.CIS3D H1lb_Hinbn 
            Height          =   360
            Left            =   120
            Top             =   450
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
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
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   360
            Left            =   120
            TabIndex        =   3
            Top             =   450
            Width           =   3195
            _ExtentX        =   5636
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
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   960
         Left            =   1365
         Top             =   60
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   1693
         ForeColor       =   16711680
         Caption         =   "仕 入 先"
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
         cPositionY      =   60
         Begin Cis3D_v60.CIS3D H1lb_Torcd 
            Height          =   345
            Left            =   75
            Top             =   450
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   609
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
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   60
            TabIndex        =   1
            Top             =   450
            Width           =   990
            _ExtentX        =   1746
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
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1050
            Top             =   450
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D15 
         Height          =   960
         Left            =   5010
         Top             =   60
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   1693
         ForeColor       =   16711680
         Caption         =   "受入"
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
         cPositionY      =   60
         Begin Cis3D_v60.CIS3D H1lb_Ukeir 
            Height          =   345
            Left            =   150
            Top             =   465
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   609
            BackColor       =   16777152
            Caption         =   "XX"
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
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   360
            Left            =   150
            TabIndex        =   2
            Top             =   450
            Width           =   390
            _ExtentX        =   688
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
            Text            =   "XX"
            MaxLength       =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   960
         Index           =   5
         Left            =   9090
         Top             =   60
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   1693
         ForeColor       =   16711680
         Caption         =   "表示年月"
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
         cPositionX      =   30
         cPositionY      =   60
         Begin CisYM_V60.CisYM iH1_Ym 
            Height          =   360
            Left            =   90
            TabIndex        =   4
            Top             =   480
            Width           =   1395
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
      Caption         =   "【 かんばん管理入力(増減・平準化)  】"
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
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   16711935
         Caption         =   "XXXXXX"
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
         Left            =   11835
         Top             =   30
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   556
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cBoderStyle     =   1
         cPositionX      =   50
      End
   End
   Begin VB.Timer Timer2 
      Left            =   540
      Top             =   60
   End
   Begin VB.Timer Timer1 
      Left            =   60
      Top             =   60
   End
   Begin VB.Timer Timer3 
      Left            =   1080
      Top             =   60
   End
End
Attribute VB_Name = "CKH0170"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   かんばん管理入力（増減・平準化 / 新Ver）
'**       フォームID    :   CKH0170
'**       処理概要      :
'**
'**       作  成  日    :   2006/02/17  By CIS
'**       変  更  日    :   2006/03/16  By CIS 増減結果(平準化)の計算を回転枚数(有効枚数)に変更
'**                         2006/03/20  By CIS 【無効】表示関連を修正
'**                         2006/03/21  By CIS 平準化参照計算時,ｶﾚﾝﾀﾞﾏｽﾀ読込み場所変更
'**                                            平準化参照計算時,回転枚数計算に累積枚数が加味されていない
'**********************************************************'
    Dim CisFrm                  As CISFormContorl

    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim SyoriKB                 As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM                 As String       ' 処理名称(新規･修正･削除 等)
    Dim InputNo                 As Integer      ' 入力番号
    Dim InputKB                 As String       ' 入力区分
    Dim InputNM                 As String       ' 入力名称
    Dim CisVsGrid1              As New CisVsGrid3
    Dim CisVsGrid2              As New CisVsGrid3

    Dim Kn_Jisya                As String
    Dim Kn_JisyaNM              As String
    Dim Kn_JisyaKJNM            As String

'   退避領域
    Dim Old_SKYmd               As String       ' 開始日（修正前）
    Dim New_EndNo               As Integer      ' 最終発行連番
    
    Dim mRType                  As Integer
    Dim mMntRow                 As Integer

    Dim mFirst_Flg              As Byte
    Dim mFirst_Gyo              As Integer

    Dim mBefore_ZGKB            As Byte

'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    If RTrim(Rv_Call2) = "" Then
    '   #-------------------#
    '   # 二 重 起 動 防 止  #
    '   #-------------------#
        If Not Dupli_Start Then End
    '   #-------------------#
    '   # 初 期 内 容 取 得  #
    '   #-------------------#
        If Not IniGet Then End
    End If
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
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
    Timer3.Interval = 500    ' 入力区分点滅間隔
    Timer3.Enabled = True    ' 入力区分点滅禁止

    Select Case App.Title
        Case "PKH0150"
            mRType = 1
'        Case "PKH0140"
'            mRType = 2
        Case Else
            mRType = 0
    End Select
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
'   かんばん増減/平準管理
    With CisVsGrid1
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 300
        .PatanMax = 1
        .SelectCol = 12
        .InitGet = False
        .Init
    End With
    Call GridInit1
'    vsGrid1.HighLight = flexHighlightNever
'   内示情報
    With CisVsGrid2
        Set .GridObj = vsGrid2
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 300
        .PatanMax = 1
        .SelectCol = 9
        .InitGet = False
        .Init
    End With
    Call GridInit2

'   #------------------#
'   #    初期入力設定  #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = "LOC"
    Call SyoriName(SyoriKB)
'   #-------------------#
'   # 画面中央位置表示  #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    If RTrim(Rv_Call2) = "" Then Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

    B1Op_ZGKbn(0) = True

' 自社情報取得
    Kn_Jisya = ""
    Kn_JisyaNM = ""
    Kn_JisyaKJNM = ""
    
    Kn_Jisya = RTrim(gCompany)
' 自社伝票発行用(工場)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 7
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 7
        IRN.oyian = "0"
        IRN.oyian2 = "00"
    End If
    ' 工場
    Kn_Jisya = Kn_Jisya & RTrim(IRN.oyian)
    
    gSL_Select = "Select * From 取引先マスタ "
    gSL_Select = gSL_Select & " Where 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & Kn_Jisya & "'"
    If TRMRead(gSL_Select, 1) Then
       Kn_JisyaNM = TRM.略称
       Kn_JisyaKJNM = TRM.工場名
    End If

    H1lb_Seban.Visible = False
    H1lb_Torcd.Visible = False
    H1lb_Ukeir.Visible = False
    H1lb_Hinbn.Visible = False
    
    If RTrim(Rv_Call2) <> "" Then
        Rv_CallRtn = False
        mFirst_Flg = 1
        
        iH1_Seban = RTrim(RV_Seban): H1lb_Seban = RTrim(RV_Seban): H1lb_Seban.Visible = True
        iH1_Torcd = RTrim(RV_Torcd): H1lb_Torcd = RTrim(RV_Torcd): H1lb_Torcd.Visible = True
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RTrim(RV_Ukeir): H1lb_Ukeir = RTrim(RV_Ukeir): H1lb_Ukeir.Visible = True
        iH1_Hinbn = RTrim(RV_Hinbn): H1lb_Hinbn = RTrim(RV_ZG_DSPHin): H1lb_Hinbn.Visible = True
        If RTrim(H1lb_Hinbn) = "" Then
           H1lb_Hinbn = RTrim(RV_Hinbn)
        End If
        iH1_Ym = RTrim(RV_YM)
        Call PB_ENT_Click
    
    End If
    
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
           Case vbKeyF1:        If PB_ADD.Visible Then Call PB_ADD_Click        '【新規】
           Case vbKeyF2:        If PB_MNT.Visible Then Call PB_MNT_Click        '【修正】
           Case vbKeyF3:        If PB_DEL.Visible Then Call PB_DEL_Click        '【削除】
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click      '【検索】
           Case vbKeyPageUp:    If PB_BACK.Visible Then Call PB_Back_Click      '【後退】
           Case vbKeyPageDown:  If PB_NEXT.Visible Then Call PB_Next_Click      '【前進】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_Can_Click        '【取消】
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click        '【終了】
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
    If RTrim(Rv_Call2) = "" Then
        Call CisDB.DBDISConnect
        Call CisFrm.UnLoadDisp
        End
    Else
        If Rv_CallRtn Then
           Call Grid_Item_Set
        End If
        Me.Hide
    End If
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
Private Sub PB_Can_Click()
    Call CisFrm.MousePT(11)
    If ProcHB = "B1" Then
        If SyoriKB = "ADD" Then
            ProcHB = "H1"
            Call DispChange(ProcHB)
            Call HeadBodyClear("B")
'            Call InputName(InputKB)
            SyoriKB = "LOC"
            Call SyoriName(SyoriKB)
            If RTrim(Rv_Call2) = "" Then iH1_Seban.SetFocus
            GoTo PB_CAN_Ed
        Else
            Call CisVsGrid1.vsColor(消去全部)
            ProcHB = "B4"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
            vsGrid1.SetFocus
            GoTo PB_CAN_Ed
        End If
    End If
    If ProcHB = "B4" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
'        Call InputName(InputKB)
        SyoriKB = "LOC"
        Call SyoriName(SyoriKB)
        If RTrim(Rv_Call2) = "" Then iH1_Seban.SetFocus
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    新  規  キ ー(F1)       +
'+----------------------------+
Private Sub PB_ADD_Click()
    If ProcHB <> "H1" Then
        If SyoriKB <> "ADD" Then
            SyoriKB = "ADD"
            ProcHB = "B1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
            Call HeadBodyClear("B", ProcHB)
            Call BodySet("ALL")
        End If
        B1Chk_SKbn.SetFocus
    End If
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If ProcHB <> "H1" Then
        If SyoriKB <> "MNT" Then
            SyoriKB = "MNT"
            ProcHB = "B4"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
        End If
        vsGrid1.SetFocus: Call BodySet1
    End If
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If ProcHB <> "H1" Then
        If SyoriKB <> "DEL" Then
            SyoriKB = "DEL"
            ProcHB = "B4"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
        End If
        vsGrid1.SetFocus: Call BodySet1
    End If
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    If ProcHB = "H1" Then
        Call CisFrm.MousePT(11)

        Call SqlSelectSet("Back")
        Call DBInput("Back")

        iH1_Seban.SetFocus

        Call CisFrm.MousePT(1)
    End If
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    If ProcHB = "H1" Then
        Call CisFrm.MousePT(11)

        Call SqlSelectSet("Next")
        Call DBInput("Next")

        iH1_Seban.SetFocus

        Call CisFrm.MousePT(1)
    End If
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Seban" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_Ukeir: Exit Sub
'   日付
    If PB_Look.Tag = "iB1_SKYmd" Then iB1_SKYmd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iB1_EKYmd" Then iB1_EKYmd.ShowCalender: Exit Sub
End Sub
'+----------------------------+
'+      品番検索              +
'+----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
    RVI_Hinmoku = 1
    RVI_Torcd = iH1_Torcd
    RVI_Ukeir = iH1_Ukeir
    RV_Call = ""

    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iH1_Seban = RV_Seban
        iH1_Hinbn = RV_Hinbn
        iH1_Torcd = RV_TorcdK
        iH1_Ukeir = RV_Ukeir
        H1lb_Tornm = ""
        If TorNmGet(iH1_Torcd, 9) Then
           H1lb_Tornm = TRM_RName
        End If
    End If
End Sub
'+----------------------------+
'+      取引先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 9            ' [発注]手配先/契約先/社内検索
    RV_Call = "C"

    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'+----------------------------+
'+      取引先受入検索        +
'+----------------------------+
Private Sub Look_Ukeir()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 9            ' [発注]手配先/契約先/社内検索
    RV_Torcd = iH1_Torcd
    RV_Kojyo = ""
    RV_Call = "C"

    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
        iH1_Ukeir = RV_Ukeir
    End If
End Sub
'+-----------------------------+
'+  iH1_Seban
'+-----------------------------+
Private Sub iH1_Seban_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Seban_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
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
    If TorNmGet(iH1_Torcd, 9) Then
       H1lb_Tornm = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  iH1_Ukeir
'+-----------------------------+
Private Sub iH1_Ukeir_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Ukeir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
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
'+-----------------------------+
'+  B1Chk_SKbn
'+-----------------------------+
Private Sub B1Chk_SKbn_Click()
    Select Case B1Chk_SKbn
        Case 0: B1Chk_SKbn.Caption = "有効"
        Case 1: B1Chk_SKbn.Caption = "無効"
    End Select
End Sub
'+-----------------------------+
'+  B1Op_ZGKbn
'+-----------------------------+
Private Sub B1Op_ZGKbn_Click(Index As Integer)
    If Index = 1 Then
       iB1_SetSu.Enabled = True
       iB1_BinSu.Enabled = True
       B1lb_AfterHeiSu = Format(NewHeiSu, "###,###")
    End If
    If Index = 2 Then
       iB1_SetSu.Enabled = True
       iB1_BinSu.Enabled = False
       iB1_BinSu = ""
       B1lb_AfterHeiSu = Format(NewHeiSu, "###,###")
    End If
    If Index = 0 Then
        iB1_BinSu.Enabled = False
        iB1_BinSu = ""
        iB1_SetSu.Enabled = False
        iB1_SetSu = ""
        B1lb_AfterHeiSu = ""
    End If
End Sub
'+-----------------------------+
'+  B1Op_HJKbn
'+-----------------------------+
Private Sub B1Op_HJKbn_Click(Index As Integer)
    If Index = 0 Then
        iB1_MonSu.Enabled = False
        iB1_MonSu = ""
        B1lb_AfterHeiSu = ""
    Else
        iB1_MonSu.Enabled = True
        B1lb_AfterHeiSu = Format(NewHeiSu, "###,###")
    End If
End Sub
'+-----------------------------+
'+  iB1_SKYmd
'+-----------------------------+
Private Sub iB1_SKYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_SKYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_EKYmd
'+-----------------------------+
Private Sub iB1_EKYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_EKYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
Private Sub iB1_SetSu_LostFocus()
    B1lb_AfterHeiSu = Format(NewHeiSu, "###,###")
End Sub
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
    Static w_Color      As Long
    w_Color = &H80000008 '初期色ｾｯﾄ(黒)
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
        Case Else
            SyoriNM = ""
    End Select
    For Each gObj In Me.Controls
        If gObj.Name Like "i*" Then gObj.ForeColor = w_Color
        If gObj.Name Like "*Comb*" Then gObj.ForeColor = w_Color
'       * PF-Key 表示設定 *
        If SYKB <> "LOC" Then
            If gObj.Name = "PB_ADD" Or _
               gObj.Name = "PB_MNT" Or _
               gObj.Name = "PB_DEL" Or _
               gObj.Name = "PB_LOK" Then
               gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
            End If
        End If
    Next gObj
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
    Dim B4Mode      As Boolean
    Dim B4Color     As Long

    H1Mode = False: H1Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
    B4Mode = False: B4Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
        Case "H1": H1Mode = True: H1Color = gPGotSel
        Case "B1": B1Mode = True: B1Color = gPGotSel
        Case "B4": B4Mode = True: B4Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode: H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode: B1_Area1.BackColor = B1Color
    B1_Area2.Enabled = B1Mode: B1_Area2.BackColor = B1Color
    B1_Area3.Enabled = B1Mode: B1_Area3.BackColor = B1Color
    B1_Area4.Enabled = B4Mode: B1_Area4.BackColor = B4Color

    PB_ADD.Visible = False
    PB_MNT.Visible = False
    PB_DEL.Visible = False
    If RTrim(Rv_Call2) = "" Then
       PB_NEXT.Visible = H1Mode
       PB_BACK.Visible = H1Mode
    Else
       PB_NEXT.Visible = False
       PB_BACK.Visible = False
    End If
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode

    If RTrim(Rv_Call2) = "" Then
        PB_END.BtnCaption = "終了"
    Else
        PB_END.BtnCaption = "戻る"
        PB_END.Visible = True
    End If
    
    If RTrim(Rv_Call2) <> "" Then
       iH1_Seban.Enabled = False: iH1_Seban.cLostColor = &HFFFFC0
       iH1_Torcd.Enabled = False: iH1_Seban.cLostColor = &HFFFFC0
       iH1_Ukeir.Enabled = False: iH1_Seban.cLostColor = &HFFFFC0
       iH1_Hinbn.Enabled = False: iH1_Seban.cLostColor = &HFFFFC0
    End If
    
'    vsGrid1.TabStop = T1Mode
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
            If .Name Like "i" & wbc_Enm & "_*" Then .Text = ""
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
        End With
    Next gObj
    B1Op_ZGKbn(0) = True
    B1Op_HJKbn(0) = True
    If HeadBody = "B" And ClrType = "ALL" Then
        Call CisVsGrid1.Clear
        Call CisVsGrid2.Clear
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        SyoriKB = "ADD"
        '初回のみ
        If RTrim(Rv_Call2) <> "" And mFirst_Flg = 1 And mFirst_Gyo <> 0 Then SyoriKB = "MNT"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        
        '初回のみ
        If RTrim(Rv_Call2) <> "" And mFirst_Flg = 1 And mFirst_Gyo <> 0 Then
           vsGrid1.Row = mFirst_Gyo
           Call vsGrid1_Click
           mFirst_Flg = 0: mFirst_Gyo = 0
        End If
        
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
'        Call HeadBodyClear("B", ProcHB)
        Call GridSet1
                
        If SyoriKB <> "ADD" Then
            ProcHB = "B4"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
            vsGrid1.SetFocus
            
        End If
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B4" Then
       Call vsGrid1_Click
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False

'   背番号
    If Trim(iH1_Seban) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    背番号を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       If iH1_Seban.Enabled = True Then iH1_Seban.SetFocus
       Exit Function
    End If

    gSL_Select = "select * from 品番取引先マスタ"
    gSL_Select = gSL_Select & " where 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   and 取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   and 背番号 = '" & RTrim(iH1_Seban) & "'"
    gSL_Select = gSL_Select & "   and 手配区分 <> '2'"
    If Not HTMRead(gSL_Select, 1, 1) Then

        If RTrim(iH1_Hinbn) <> "" And RTrim(iH1_Torcd) <> "" And RTrim(iH1_Ukeir) <> "" And RTrim(iH1_Seban) <> "" Then
           With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    品番取引先マスタ未登録            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
           End With
           If iH1_Seban.Enabled = True Then iH1_Seban.SetFocus
           Exit Function
        End If
        
        gSL_Select = ""
        gSL_Select = gSL_Select & "Select"
        gSL_Select = gSL_Select & "  htm.背番号"
        gSL_Select = gSL_Select & ", htm.取引先"
        gSL_Select = gSL_Select & ", htm.受入"
        gSL_Select = gSL_Select & ", htm.品番"
        gSL_Select = gSL_Select & ", hnm.表示品番"
        gSL_Select = gSL_Select & ", hnm.社内背番"
        
        gSL_Select = gSL_Select & ", Case When htm.最終発行連番 > KN.発行NO"
        gSL_Select = gSL_Select & "  Then htm.最終発行連番 Else KN.発行NO  End 最終発行連番"
        
        gSL_Select = gSL_Select & ", rtrim(trm.略称) 取引先名"
        gSL_Select = gSL_Select & ", htm.手配区分"
        gSL_Select = gSL_Select & " from 品番取引先マスタ htm"
        gSL_Select = gSL_Select & "      left outer join 品番マスタ hnm"
        gSL_Select = gSL_Select & "        on hnm.品番 = htm.品番"
        gSL_Select = gSL_Select & "      left outer join 取引先マスタ trm"
        gSL_Select = gSL_Select & "        on trm.取引先CD = htm.取引先"
        gSL_Select = gSL_Select & "       and trm.取引先区分 = 1"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN ( SELECT 取引先, 受入, 品番, MAX(発行NO) 発行NO"
        gSL_Select = gSL_Select & " From かんばん流動管理テーブル"
        gSL_Select = gSL_Select & " GROUP BY 取引先, 受入, 品番 ) KN"
        gSL_Select = gSL_Select & " ON  KN.取引先 = htm.取引先"
        gSL_Select = gSL_Select & " AND KN.受入 = htm.受入"
        gSL_Select = gSL_Select & " AND KN.品番 = htm.品番"
        gSL_Select = gSL_Select & " where htm.背番号 = '" & iH1_Seban & "'"
        gSL_Select = gSL_Select & "   and htm.品目 = 1"
        gSL_Select = gSL_Select & "   and htm.SK区分 = 0"
        gSL_Select = gSL_Select & "   and htm.手配区分 <> '2'"
        gSL_Select = gSL_Select & " order by"
        gSL_Select = gSL_Select & "  htm.品番"
        gSL_Select = gSL_Select & ", htm.取引先"
        gSL_Select = gSL_Select & ", htm.受入"

'
        If HTMRead(gSL_Select, 0, 1) Then
            If CisDB.RecordCount > 1 Then
                CKN0215.Show vbModal
                Unload CKN0215
                Set CKN0215 = Nothing
                If RV_Rtn Then
                    iH1_Torcd = RV_TorcdK
                    iH1_Ukeir = RV_Ukeir
                    iH1_Hinbn = RV_Hinbn
                    Call iH1_Torcd_LostFocus
                Else
                    iH1_Seban.SetFocus
                    Exit Function
                End If
            Else
                iH1_Torcd = RTrim(HTM.取引先)
                iH1_Ukeir = RTrim(HTM.受入)
                iH1_Hinbn = RTrim(HTM.品番)
            End If
            Call HTMClose(1)
        Else
            Call HTMClose(1)

'           取引先
            If Trim(iH1_Torcd) = "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    仕入先を入力して下さい            "
                    .MB_Button = Error
                    .MBOX
                End With
                If iH1_Torcd.Enabled = True Then iH1_Torcd.SetFocus
                Exit Function
            End If
            If Not TorNmGet(iH1_Torcd, 9) Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    仕入先　未登録！（取引先マスタ）            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
               If iH1_Torcd.Enabled = True Then iH1_Torcd.SetFocus
               Exit Function
            End If
'           受入
            If Trim(iH1_Ukeir) = "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    受入を入力して下さい            "
                    .MB_Button = Error
                    .MBOX
                End With
                If iH1_Ukeir.Enabled = True Then iH1_Ukeir.SetFocus
                Exit Function
            End If
            gSL_Select = ""
            gSL_Select = gSL_Select & "select *"
            gSL_Select = gSL_Select & "  from 取引先受入マスタ"
            gSL_Select = gSL_Select & " where 取引先CD = '" & RTrim(iH1_Torcd) & "'"
            gSL_Select = gSL_Select & "   and 取引先区分 = 1"
            gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(iH1_Ukeir) & "'"
            If Not TUMRead(gSL_Select, 1) Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    受入　未登録！（取引先受入マスタ）            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
               If iH1_Ukeir.Enabled = True Then iH1_Ukeir.SetFocus
               Exit Function
            End If
'           品番
            If Trim(iH1_Hinbn) = "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    品番を入力して下さい            "
                    .MB_Button = Error
                    .MBOX
                End With
                If iH1_Hinbn.Enabled = True Then iH1_Hinbn.SetFocus
                Exit Function
            End If
'           <<< 品番取引先マスタ >>>
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    品番取引先マスタ未登録            "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
            If iH1_Seban.Enabled = True Then iH1_Seban.SetFocus
            Exit Function

        End If
    End If

'   年月
'    If Trim(iH1_Ym) = "" Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    年月を入力して下さい            "
'            .MB_Title = ""
'            .MB_Button = Error
'            .MBOX
'        End With
'       iH1_Ym.SetFocus
'       Exit Function
'    End If
    
    If Trim(iH1_Ym) <> "" Then
        If iH1_Ym.cChkResult = False Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    正しい年月を入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_Ym.SetFocus
           Exit Function
        End If
    End If
    Call SqlSelectSet("INV")
    If Not DBInput("INV") Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    該当データが存在しません｡            "
            .MB_MSG(3) = "    再度抽出条件を入力してください       "
            .MB_Title = "抽出処理"
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If

    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Dim wSKYmd      As String
    Dim wEKYmd      As String

    Body1Chk = False

    If SyoriKB = "MNT" Then
       If mMntRow = 0 Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    修正する行を選択して下さい        "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          vsGrid1.SetFocus
          Exit Function
       End If
    End If
'   開始日
    If Trim(iB1_SKYmd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    開始日を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_SKYmd.SetFocus
       Exit Function
    Else
        If Not iB1_SKYmd.cChkResult Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    開始日を正しく入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB1_SKYmd.SetFocus
            Exit Function
        End If
    End If
'   終了日
    If Trim(iB1_EKYmd) <> "" Then
        If Not iB1_EKYmd.cChkResult Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    終了日を正しく入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB1_EKYmd.SetFocus
            Exit Function
        End If
    '   大小関係
        If iB1_SKYmd > iB1_EKYmd Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    日付を正しく入力して下さい（開始＞終了）            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB1_SKYmd.SetFocus
            Exit Function
        End If
    End If
    
    If B1Op_ZGKbn(0).Value = True And B1Op_HJKbn(0).Value = True Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    増減・平準の設定を有効にして下さい      "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        B1Op_ZGKbn(0).SetFocus
        Exit Function
    End If

    If B1Op_ZGKbn(0).Value = False Then
       If CisFun.Val2(iB1_SetSu) = 0 Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    増減枚数を入力して下さい      "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB1_SetSu.SetFocus
           Exit Function
       End If
       
       If CisFun.Val2(iB1_SetSu) < CisFun.Val2(iB1_BinSu) Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    便枚数が設定枚数をオーバーしています     "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB1_BinSu.SetFocus
           Exit Function
       End If
       
    End If
    If B1Op_HJKbn(0).Value = False Then
       If CisFun.Val2(iB1_MonSu) = 0 Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    平準枚数を入力して下さい      "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB1_MonSu.SetFocus
           Exit Function
       End If
    End If
    
'   <><><><><> 増減区分 = 設定なし以外 <><><><><>
    If Not B1Op_ZGKbn(0).Value Then
    '   <><><><><> 重複ﾁｪｯｸ／設定：有効の時 <><><><><>
'        If B1Chk_SKbn = 0 Then
            wSKYmd = iB1_SKYmd
            wEKYmd = iB1_EKYmd
            If Trim(iB1_EKYmd) = "" Then wEKYmd = "99991231"
            gSL_Select = ""
            gSL_Select = gSL_Select & "select kzt.開始日"
            gSL_Select = gSL_Select & "     , kzt.終了日"
            gSL_Select = gSL_Select & "  from かんばん増減テーブル kzt"
            gSL_Select = gSL_Select & " where kzt.取引先 = '" & RTrim(iH1_Torcd) & "'"
            gSL_Select = gSL_Select & "   and kzt.受入 = '" & RTrim(iH1_Ukeir) & "'"
            gSL_Select = gSL_Select & "   and kzt.品番 = '" & RTrim(iH1_Hinbn) & "'"
            If SyoriKB = "MNT" Then
                gSL_Select = gSL_Select & "   and kzt.開始日 <> '" & RTrim(Old_SKYmd) & "'"
            End If
'            gSL_Select = gSL_Select & "   and kzt.設定内容区分 = 0"
            gSL_Select = gSL_Select & "   and kzt.開始日 <= '" & wEKYmd & "'"
            gSL_Select = gSL_Select & "   and case when kzt.終了日 = ''"
            gSL_Select = gSL_Select & "            then '99991231'"
            gSL_Select = gSL_Select & "            else kzt.終了日"
            gSL_Select = gSL_Select & "        end >= '" & wSKYmd & "'"
            If KZTRead(gSL_Select, 1) Then
                With CisFun
                    .MB_Lines = 5
                    .MB_MSG(2) = "    下記の登録済日付と重複しています！            "
                    .MB_MSG(4) = "        > 開始日 : " & Format(KZT.開始日, "0000/00/00")
                    .MB_MSG(5) = "        > 終了日 : " & Format(KZT.終了日, "0000/00/00")
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                iB1_SKYmd.SetFocus
                Exit Function
            End If
'        End If
    
    '   設定枚数
        If CisFun.Val2(iB1_SetSu) = 0 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    設定枚数を入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iB1_SetSu.SetFocus
           Exit Function
        End If
    
    End If

'   <><><><><> 平準区分 = 設定なし以外 <><><><><>
    If Not B1Op_HJKbn(0).Value Then
    '   <><><><><> 重複ﾁｪｯｸ／設定：有効の時 <><><><><>
'        If B1Chk_SKbn = 0 Then
            wSKYmd = iB1_SKYmd
            wEKYmd = iB1_EKYmd
            If Trim(iB1_EKYmd) = "" Then wEKYmd = "99991231"
            gSL_Select = ""
            gSL_Select = gSL_Select & "select kjt.開始日"
            gSL_Select = gSL_Select & "     , kjt.終了日"
            gSL_Select = gSL_Select & "  from かんばん平準化テーブル kjt"
            gSL_Select = gSL_Select & " where kjt.取引先 = '" & RTrim(iH1_Torcd) & "'"
            gSL_Select = gSL_Select & "   and kjt.受入 = '" & RTrim(iH1_Ukeir) & "'"
            gSL_Select = gSL_Select & "   and kjt.品番 = '" & RTrim(iH1_Hinbn) & "'"
            If SyoriKB = "MNT" Then
                gSL_Select = gSL_Select & "   and kjt.開始日 <> '" & RTrim(Old_SKYmd) & "'"
            End If
'            gSL_Select = gSL_Select & "   and kjt.設定内容区分 = 0"
            gSL_Select = gSL_Select & "   and kjt.開始日 <= '" & wEKYmd & "'"
            gSL_Select = gSL_Select & "   and case when kjt.終了日 = ''"
            gSL_Select = gSL_Select & "            then '99991231'"
            gSL_Select = gSL_Select & "            else kjt.終了日"
            gSL_Select = gSL_Select & "        end >= '" & wSKYmd & "'"
            If KJTRead(gSL_Select, 1) Then
                With CisFun
                    .MB_Lines = 5
                    .MB_MSG(2) = "    下記の登録済日付と重複しています！            "
                    .MB_MSG(4) = "        > 開始日 : " & Format(KJT.開始日, "0000/00/00")
                    .MB_MSG(5) = "        > 終了日 : " & Format(KJT.終了日, "0000/00/00")
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                iB1_SKYmd.SetFocus
                Exit Function
            End If
'        End If
    '   上限枚数
        If CisFun.Val2(iB1_MonSu) = 0 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    上限枚数を入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iB1_MonSu.SetFocus
           Exit Function
        End If
    
    
        B1lb_AfterHeiSu = Format(NewHeiSu, "###,###")
    End If
    
    Body1Chk = True
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    
    Call HeadBodyClear("B")
    If Not HTMRead(gSL_Select, 1) Then
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
        iH1_Seban = RTrim(HTM.背番号)
        iH1_Hinbn = RTrim(HTM.品番)
        iH1_Torcd = RTrim(HTM.取引先)
        iH1_Ukeir = RTrim(HTM.受入)
        H1lb_Tornm = RTrim(HTM.取引先名)
        If InputKey = "INV" Then
            If Not BodySet("ALL") Then GoTo DBInput_Ed
        End If
        DBInput = True
    End If

DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'****************************
'*      SQL文ｾｯﾄ処理
'****************************
Private Sub SqlSelectSet(InputKey As String)
'
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  htm.品番"
    gSL_Select = gSL_Select & ", htm.取引先"
    gSL_Select = gSL_Select & ", htm.受入"
    gSL_Select = gSL_Select & ", htm.背番号"
    gSL_Select = gSL_Select & ", htm.か係数"
    gSL_Select = gSL_Select & ", htm.サイクル"
    gSL_Select = gSL_Select & ", htm.収容数1"
    gSL_Select = gSL_Select & ", htm.収容器1"
    gSL_Select = gSL_Select & ", htm.変更日"
    gSL_Select = gSL_Select & ", htm.変更便"
    gSL_Select = gSL_Select & ", htm.収容数2"
    gSL_Select = gSL_Select & ", htm.収容器2"
    gSL_Select = gSL_Select & ", htm.平準化区分"
    gSL_Select = gSL_Select & ", htm.平準化係数"
    gSL_Select = gSL_Select & ", htm.置場"
    
    gSL_Select = gSL_Select & ", Case When htm.最終発行連番 > KN.発行NO"
    gSL_Select = gSL_Select & "  Then htm.最終発行連番 Else KN.発行NO  End 最終発行連番"
    
    gSL_Select = gSL_Select & ", isnull(trm.略称, '') 取引先名"
    gSL_Select = gSL_Select & ", isnull(hnm.社内背番, '') 社内背番"
    gSL_Select = gSL_Select & ", isnull(hnm.品名, '') 品名"
    gSL_Select = gSL_Select & ", isnull(hnm.車型, '') 車型"
    gSL_Select = gSL_Select & ", isnull(sym.値名称, '') 手配区分名"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  品番取引先マスタ htm"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ trm"
    gSL_Select = gSL_Select & "    on trm.取引先CD = htm.取引先"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 1"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN ( SELECT 取引先, 受入, 品番, MAX(発行NO) 発行NO"
    gSL_Select = gSL_Select & " From かんばん流動管理テーブル"
    gSL_Select = gSL_Select & " GROUP BY 取引先, 受入, 品番 ) KN"
    gSL_Select = gSL_Select & " ON  KN.取引先 = htm.取引先"
    gSL_Select = gSL_Select & " AND KN.受入 = htm.受入"
    gSL_Select = gSL_Select & " AND KN.品番 = htm.品番"
    gSL_Select = gSL_Select & "  left outer join 品番マスタ hnm"
    gSL_Select = gSL_Select & "    on hnm.品番 = htm.品番"
    gSL_Select = gSL_Select & "  left outer join 名称マスタ sym"
    gSL_Select = gSL_Select & "    on sym.区分名称 = '手配区分'"
    gSL_Select = gSL_Select & "   and sym.区分タイプ = ''"
    gSL_Select = gSL_Select & "   and sym.値 = htm.手配区分"
    gSL_Select = gSL_Select & " where htm.品目 = 1"
    gSL_Select = gSL_Select & "   and htm.SK区分 = 0"
    gSL_Select = gSL_Select & "   and htm.手配区分 <> '2'"
    gSL_Select = gSL_Select & "   and htm.背番号 <> ''"
    gSL_Select = gSL_Select & "   and"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/10 UPDATE START
    'gSL_Select = gSL_Select & " ( htm.背番号 + space(5 - len(htm.背番号))"
    gSL_Select = gSL_Select & " ( htm.背番号 + space(8 - len(htm.背番号))"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/10 UPDATE END
    gSL_Select = gSL_Select & " + htm.品番 + space(25 - len(htm.品番))"
    gSL_Select = gSL_Select & " + htm.取引先 + space(7 - len(htm.取引先))"
    gSL_Select = gSL_Select & " + htm.受入 + space(2 - len(htm.受入))"
    gSL_Select = gSL_Select & " ) "
    Select Case InputKey
        Case "Back": gSL_Select = gSL_Select & "<"
        Case "Next": gSL_Select = gSL_Select & ">"
        Case Else:   gSL_Select = gSL_Select & "="
    End Select
    gSL_Select = gSL_Select & " ( '"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/10 UPDATE START
    'gSL_Select = gSL_Select & iH1_Seban & Space(5 - CisFun.Len2(iH1_Seban))
    gSL_Select = gSL_Select & iH1_Seban & Space(8 - CisFun.Len2(iH1_Seban))
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/10 UPDATE END
    gSL_Select = gSL_Select & iH1_Hinbn & Space(25 - CisFun.Len2(iH1_Hinbn))
    gSL_Select = gSL_Select & iH1_Torcd & Space(7 - CisFun.Len2(iH1_Torcd))
    gSL_Select = gSL_Select & iH1_Ukeir & Space(2 - CisFun.Len2(iH1_Ukeir))
    gSL_Select = gSL_Select & "' )"
    Select Case InputKey
        Case "Back"
            gSL_Select = gSL_Select & " order by"
            gSL_Select = gSL_Select & "  htm.背番号 desc"
            gSL_Select = gSL_Select & ", htm.取引先 desc"
            gSL_Select = gSL_Select & ", htm.受入 desc"
            gSL_Select = gSL_Select & ", htm.品番 desc"
        Case "Next"
            gSL_Select = gSL_Select & " order by"
            gSL_Select = gSL_Select & "  htm.背番号 asc"
            gSL_Select = gSL_Select & ", htm.取引先 asc"
            gSL_Select = gSL_Select & ", htm.受入 asc"
            gSL_Select = gSL_Select & ", htm.品番 asc"
    End Select
'
End Sub
'**************************************
'*          BODY 内容ｾｯﾄ処理           *
'**************************************
Private Function BodySet(SetType As String) As Boolean
    BodySet = False
    Dim wCurrency       As Currency

    With HTM
'       << マスタ情報 >>
        If Trim(.サイクル) <> "" Then
            B1lb_Cycle = CisFun.RSetFld(CisFun.Mid2(.サイクル, 1, 1), 1, "#") _
                 & "-" & CisFun.RSetFld(CisFun.Mid2(.サイクル, 2, 2), 2, "#") _
                 & "-" & CisFun.RSetFld(CisFun.Mid2(.サイクル, 4, 2), 2, "#")
        End If
        If .か係数 <> 0 Then
           B1lb_KKeisu = Format(.か係数, "0.00")
        Else
           B1lb_KKeisu = ""
        End If
        B1lb_HKeisu = Format(.平準化係数, "##")
        Select Case .平準化区分
               Case 1: B1lb_HKbn = "日"
               Case 2: B1lb_HKbn = "便"
               Case 3: B1lb_HKbn = "月"
               Case Else: B1lb_HKbn = "無"
        End Select
        B1lb_Okiba = RTrim(HTM.置場)
        '開始日マスタ検索
    '開始日の検索、存在しなければ基本開始日をセット
        gSL_Select = "SELECT 1 区分,* FROM かんばん増減開始日マスタ"
        gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & " AND   受入 = ''"
        gSL_Select = gSL_Select & " AND   品番 = ''"
        gSL_Select = gSL_Select & " UNION "
        gSL_Select = gSL_Select & " SELECT 2 区分,* FROM かんばん増減開始日マスタ"
        gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & " AND   受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & " AND   品番 = ''"
        gSL_Select = gSL_Select & " UNION "
        gSL_Select = gSL_Select & " SELECT 3 区分,* FROM かんばん増減開始日マスタ"
        gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & " AND   受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & " AND   品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & " ORDER BY 区分 DESC"
        If KZSRead(gSL_Select, 1, 1) Then
           B1lb_StartDay = Format(KZS.開始日, "##")
        End If
        
'       【かんばん発行】
        B1lb_Syuki(1) = RTrim(.収容器1)
        B1lb_Syuyo(1) = Format(.収容数1, "#,###")
        If Trim(.変更日) <> "" Then
            B1lb_SYmd = Format(.変更日, "0000/00/00") & " " & .変更便
            B1lb_Syuki(2) = RTrim(.収容器2)
            B1lb_Syuyo(2) = Format(.収容数2, "#,###")
        End If
        
'        B1lb_EndNo = Format(.最終発行連番, "#,###")
    End With

'   <><><> かんばん流動管理テーブル <><><>
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  sum(case when krt.発行区分 = 0 then 1 else 0 end) 未流件数"
    gSL_Select = gSL_Select & ", sum(case when krt.発行区分 = 1 then 1 else 0 end) 有効件数"
    gSL_Select = gSL_Select & ", sum(case when krt.発行区分 = 2 then 1 else 0 end) 無効件数"
    gSL_Select = gSL_Select & ", sum(case when krt.発行区分 = 3 then 1 else 0 end) 滞留件数"
    gSL_Select = gSL_Select & " from かんばん流動管理テーブル krt"
    gSL_Select = gSL_Select & " where krt.取引先 = '" & RTrim(HTM.取引先) & "'"
    gSL_Select = gSL_Select & "   and krt.受入 = '" & RTrim(HTM.受入) & "'"
    gSL_Select = gSL_Select & "   and krt.品番 = '" & RTrim(HTM.品番) & "'"
    gSL_Select = gSL_Select & "   and krt.収容区分 = 1"
    If KRTRead(gSL_Select, 1, 1) Then
        With KRT
'           << 流動管理情報 >>
            B1lb_HMaisu(1) = Format(.有効件数 + .無効件数 + .滞留件数 + .未流件数, "#,###")
            B1lb_YSu(1) = Format(.有効件数, "#,###")
            B1lb_MSu(1) = Format(.無効件数, "#,###")
            B1lb_TSu(1) = Format(.滞留件数, "#,###")
            B1lb_MISu(1) = Format(.未流件数, "#,###")
        End With
    End If
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  sum(case when krt.発行区分 = 0 then 1 else 0 end) 未流件数"
    gSL_Select = gSL_Select & ", sum(case when krt.発行区分 = 1 then 1 else 0 end) 有効件数"
    gSL_Select = gSL_Select & ", sum(case when krt.発行区分 = 2 then 1 else 0 end) 無効件数"
    gSL_Select = gSL_Select & ", sum(case when krt.発行区分 = 3 then 1 else 0 end) 滞留件数"
    gSL_Select = gSL_Select & " from かんばん流動管理テーブル krt"
    gSL_Select = gSL_Select & " where krt.取引先 = '" & RTrim(HTM.取引先) & "'"
    gSL_Select = gSL_Select & "   and krt.受入 = '" & RTrim(HTM.受入) & "'"
    gSL_Select = gSL_Select & "   and krt.品番 = '" & RTrim(HTM.品番) & "'"
    gSL_Select = gSL_Select & "   and krt.収容区分 = 2"
    If KRTRead(gSL_Select, 1, 1) Then
        With KRT
'           << 流動管理情報 >>
            B1lb_HMaisu(2) = Format(.有効件数 + .無効件数 + .滞留件数 + .未流件数, "#,###")
            B1lb_YSu(2) = Format(.有効件数, "#,###")
            B1lb_MSu(2) = Format(.無効件数, "#,###")
            B1lb_TSu(2) = Format(.滞留件数, "#,###")
            B1lb_MISu(2) = Format(.未流件数, "#,###")
        End With
    End If
    
    B1lb_HMaisu(0) = Format(CisFun.Val2(B1lb_YSu(1)) + CisFun.Val2(B1lb_YSu(2)) + CisFun.Val2(B1lb_MSu(1)) + CisFun.Val2(B1lb_MSu(2)) + CisFun.Val2(B1lb_TSu(1)) + CisFun.Val2(B1lb_TSu(2)) + CisFun.Val2(B1lb_MISu(1)) + CisFun.Val2(B1lb_MISu(2)), "#,###")
    B1lb_YSu(0) = Format(CisFun.Val2(B1lb_YSu(1)) + CisFun.Val2(B1lb_YSu(2)), "#,###")
    B1lb_MSu(0) = Format(CisFun.Val2(B1lb_MSu(1)) + CisFun.Val2(B1lb_MSu(2)), "#,###")
    B1lb_TSu(0) = Format(CisFun.Val2(B1lb_TSu(1)) + CisFun.Val2(B1lb_TSu(2)), "#,###")
    B1lb_MISu(0) = Format(CisFun.Val2(B1lb_MISu(1)) + CisFun.Val2(B1lb_MISu(2)), "#,###")
    
'   ■■■■■■■■■■■ << 内示情報 >>  ■■■■■■■■■■
    'Grid表示
    Call CisVsGrid2.Clear
    ' 内示数
    If RTrim(iH1_Ym) <> "" Then
       gSL_Select = "select "
    Else
       gSL_Select = "select Top 5 "
    End If
    gSL_Select = gSL_Select & " 内示年月,SUM(数量) 数量 from 発注内示テーブル"
    gSL_Select = gSL_Select & " Where 手配先 = '" & RTrim(HTM.取引先) & "'"
    gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(HTM.品番) & "'"
    gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(HTM.受入) & "'"
    If RTrim(iH1_Ym) <> "" Then
       gSL_Select = gSL_Select & "   and 内示年月 = 計画年月"
       gSL_Select = gSL_Select & "   and 内示年月 < '" & RTrim(iH1_Ym) & "'"
    End If
    gSL_Select = gSL_Select & "  Group by 内示年月 "
    gSL_Select = gSL_Select & " Order By 内示年月 Desc "
    gLong = 0
    If HNJRead(gSL_Select, , 1) Then
       With vsGrid2
           If CisVsGrid2.DispMax < CisDB.RecordCount + 1 Then
              .Rows = CisDB.RecordCount + 1
           End If
           Do Until Not HNJ_RDSTS
              gLong = gLong + 1
              .TextMatrix(gLong, 0) = gLong
              .TextMatrix(gLong, 1) = Mid(HNJ.内示年月, 1, 4) & "/" & Mid(HNJ.内示年月, 5, 2)
              .TextMatrix(gLong, 2) = Format(HNJ.数量, "#,###")
              '年月
              .TextMatrix(gLong, 9) = HNJ.内示年月
              If RTrim(iH1_Ym) <> "" Then
                 If gLong = 1 Then
                    B1lb_ZenNaisu = "(" & Format(HNJ.数量, "###,###,##0") & ")"
                 End If
              End If
              
              Call HNJReadNext(1)
              
           Loop
       End With
    End If
    Call HNJClose(1)
    ' 日量枚数計算
    gCnt1 = CisFun.Val2(B1lb_Syuyo(1))
    If gCnt1 <> 0 Then
        With vsGrid2
            For gLong = 1 To .Rows - 1
                If Trim(.TextMatrix(gLong, 1)) = "" Then Exit For
                
                gSL_Select = "Select * from カレンダマスタ "
                gSL_Select = gSL_Select & " Where 取引先区分 = 1"
                gSL_Select = gSL_Select & "   And 取引先 = ''"
                gSL_Select = gSL_Select & "   And 年 + 月 = '" & Trim(.TextMatrix(gLong, 9)) & "'"
                gSL_Select = gSL_Select & "   And 受入 = ''"
                If Not CLMRead(gSL_Select, 1, 1) Then
                   CLM.稼動日数 = 31
                End If
                wCurrency = CisFun.Val2(.TextMatrix(gLong, 2)) / CLM.稼動日数 / gCnt1
                gCnt2 = CisFun.HasuProc(CDbl(wCurrency), 切捨て)
                If gCnt2 = 0 And wCurrency <> 0 Then gCnt2 = 1
                .TextMatrix(gLong, 3) = Format(gCnt2, "###")
                
                If RTrim(iH1_Ym) <> "" Then
                   If gLong = 1 Then
                      B1lb_ZenDaysu = "(" & Format(gCnt2, "###") & ")"
                   End If
                End If
            Next gLong
        End With
    End If
    
    ' かんばん回転枚数テーブル
    If RTrim(iH1_Ym) <> "" Then
       gSL_Select = "select "
    Else
       gSL_Select = "select Top 5 "
    End If
    gSL_Select = gSL_Select & " * From かんばん回転枚数テーブル "
    gSL_Select = gSL_Select & " Where 仕入先 = '" & RTrim(HTM.取引先) & "'"
    gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(HTM.品番) & "'"
    gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(HTM.受入) & "'"
    gSL_Select = gSL_Select & "   and 背番号 = '" & RTrim(HTM.背番号) & "'"
    If RTrim(iH1_Ym) <> "" Then
       gSL_Select = gSL_Select & "   and 年月 < '" & RTrim(iH1_Ym) & "'"
    End If
    gSL_Select = gSL_Select & " Order By 年月 Desc "
    gLong = 0
    If KKTRead(gSL_Select, , 1) Then
       With vsGrid2
           Do Until Not KKT_RDSTS
              For gLong = 1 To .Rows - 1
                  If Trim(.TextMatrix(gLong, 9)) = Trim(KKT.年月) Then
                    '回転枚数
                    .TextMatrix(gLong, 4) = CisFun.RSetFld(KKT.回転枚数, 3, "##0") & _
                                          "(" & CisFun.RSetFld(KKT.前回回転枚数, 3, "##0") & ")"
                    '前月差
                    .TextMatrix(gLong, 5) = Format(KKT.回転枚数 - KKT.前月回転枚数, "##0")
                    '平準化日
                    .TextMatrix(gLong, 6) = Format(KKT.日上限枚数, "###")
                    '平準化便
                    .TextMatrix(gLong, 7) = Format(KKT.便上限枚数, "###")
                    Exit For
                   End If
              Next gLong
              Call KKTReadNext(1)
          Loop
       End With
    End If
    Call KKTClose(1)

    '当月情報表示
    ' 内示数
    If RTrim(iH1_Ym) <> "" Then
        gSL_Select = "select 内示年月,SUM(数量) 数量 from 発注内示テーブル"
        gSL_Select = gSL_Select & " Where 手配先 = '" & RTrim(HTM.取引先) & "'"
        gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(HTM.品番) & "'"
        gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(HTM.受入) & "'"
        gSL_Select = gSL_Select & "   and 内示年月 = 計画年月"
        gSL_Select = gSL_Select & "   and 内示年月 = '" & RTrim(iH1_Ym) & "'"
        gSL_Select = gSL_Select & "  Group by 内示年月 "
        gSL_Select = gSL_Select & " Order By 内示年月 Desc "
        gLong = 0
        If HNJRead(gSL_Select, 1, 1) Then
           B1lb_NYM = Mid(HNJ.内示年月, 1, 4) & "年" & Mid(HNJ.内示年月, 5, 2) & "月"
           B1lb_Naisu = Format(HNJ.数量, "###,###,###")
        End If
        ' 日量枚数計算
        gCnt1 = CisFun.Val2(B1lb_Syuyo(1))
        gSL_Select = "Select * from カレンダマスタ "
        gSL_Select = gSL_Select & " Where 取引先区分 = 1"
        gSL_Select = gSL_Select & "   And 取引先 = ''"
        gSL_Select = gSL_Select & "   And 年 + 月 = '" & Trim(HNJ.内示年月) & "'"
        gSL_Select = gSL_Select & "   And 受入 = ''"
        If Not CLMRead(gSL_Select, 1, 1) Then
           CLM.稼動日数 = 31
        End If
        wCurrency = CisFun.Val2(B1lb_Naisu) / CLM.稼動日数 / gCnt1
        gCnt2 = CisFun.HasuProc(CDbl(wCurrency), 切捨て)
        If gCnt2 = 0 And wCurrency <> 0 Then gCnt2 = 1
        B1lb_Daysu = Format(gCnt2, "###")
        ' かんばん回転枚数テーブル
        gSL_Select = "Select * From かんばん回転枚数テーブル "
        gSL_Select = gSL_Select & " Where 仕入先 = '" & RTrim(HTM.取引先) & "'"
        gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(HTM.品番) & "'"
        gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(HTM.受入) & "'"
        gSL_Select = gSL_Select & "   and 背番号 = '" & RTrim(HTM.背番号) & "'"
        gSL_Select = gSL_Select & "   and 年月 = '" & RTrim(iH1_Ym) & "'"
        gSL_Select = gSL_Select & " Order By 年月 Desc "
        If KKTRead(gSL_Select, , 1) Then
            '回転枚数
            B1lb_KaiMai = CisFun.RSetFld(KKT.回転枚数, 3, "###")
            '前月
            B1lb_ZenKaiMai = "(" & Format(KKT.前月回転枚数, "###") & ")"
            '前月差
            B1lb_KaiMaiSa = Format(KKT.回転枚数 - KKT.前月回転枚数, "###")
        End If
    End If
    
'   【かんばん増減/平準管理】
    Call GridSet1

    BodySet = True
End Function
'+-------------------------------------+
'+      ｸﾞﾘｯﾄﾞより内容セット後入力処理
'+-------------------------------------+
Private Sub FromGrid1()
    If SyoriKB = "MNT" Or _
       SyoriKB = "DEL" Then
        Call BodySet1
        If Trim(vsGrid1.TextMatrix(vsGrid1.Row, CisVsGrid1.FixedGet(0))) = "" Then
            vsGrid1.SetFocus
            Exit Sub
        End If
        If SyoriKB = "MNT" Then
            Call CisVsGrid1.vsColor(消去全部)
            Call CisVsGrid1.vsColor
'            B1Chk_SKbn.SetFocus
        End If
        If SyoriKB = "DEL" Then
            Call CisVsGrid1.vsColor
            If Not DBDelete Then
                Call CisVsGrid1.vsColor(消去全部)
                vsGrid1.SetFocus
                Exit Sub
            Else
                If Not GridSet1 Then
                '   表示明細が無い時は、新規にする
                    SyoriKB = "ADD"
                    Call SyoriName(SyoriKB)
                    ProcHB = "B1"
                    Call DispChange(ProcHB)
                    B1Chk_SKbn.SetFocus
                    Exit Sub
                Else
                    ProcHB = "B4"
                    Call DispChange(ProcHB)
                    vsGrid1.SetFocus
                    Exit Sub
                End If
                Call BodySet1
            End If
            vsGrid1.SetFocus
        End If
    
        ProcHB = "B1"
        Call DispChange(ProcHB)
    
    End If
End Sub
'+---------------------------+
'+      ｸﾞﾘｯﾄﾞより内容セット
'+---------------------------+
Private Sub BodySet1()
    With vsGrid1
        If .Row < 1 Then Exit Sub
        Select Case .TextMatrix(.Row, CisVsGrid1.FixedGet(1))
            Case "有効": B1Chk_SKbn = 0
            Case "無効": B1Chk_SKbn = 1
        End Select
        iB1_SKYmd = Format(.TextMatrix(.Row, CisVsGrid1.FixedGet(2)), "yyyymmdd")
        iB1_SKYmd.Tag = iB1_SKYmd
        iB1_EKYmd = Format(.TextMatrix(.Row, CisVsGrid1.FixedGet(3)), "yyyymmdd")
        Select Case .TextMatrix(.Row, CisVsGrid1.FixedGet(4))
            Case "減": B1Op_ZGKbn(1) = True
            Case "増": B1Op_ZGKbn(2) = True
            Case Else: B1Op_ZGKbn(0) = True
        End Select
        iB1_SetSu = .TextMatrix(.Row, CisVsGrid1.FixedGet(5))
        iB1_BinSu = .TextMatrix(.Row, CisVsGrid1.FixedGet(6))
        B1lb_RuiSu = .TextMatrix(.Row, CisVsGrid1.FixedGet(7))
        B1lb_KanYmd = .TextMatrix(.Row, CisVsGrid1.FixedGet(8))
        Select Case .TextMatrix(.Row, CisVsGrid1.FixedGet(9))
            Case "日": B1Op_HJKbn(2) = True
            Case "便": B1Op_HJKbn(1) = True
            Case "月": B1Op_HJKbn(3) = True
            Case Else: B1Op_HJKbn(0) = True
        End Select
        iB1_MonSu = .TextMatrix(.Row, CisVsGrid1.FixedGet(10))
    End With
'   退避
    Old_SKYmd = iB1_SKYmd

    If SyoriKB <> "DEL" Then
       B1lb_AfterHeiSu = Format(NewHeiSu, "###,###")
    Else
       B1lb_AfterHeiSu = ""
    End If
    
End Sub
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+      削 除 処 理      +
'+----------------------+
Private Function DBDelete()
    DBDelete = False

'  ( 確認メッセージ )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    With CisFun
         .MB_Lines = 5
         .MB_MSG(2) = "    削除処理を行います。            "
         .MB_MSG(4) = "     よろしいですか？            "
         .MB_Title = "削除確認"
         .MB_Button = OK_CAN
         If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBDelete_Err

    Call CisDB.DBTran(TransBegin)   ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

' ( 削除処理 )
    If Not B1Op_ZGKbn(0) Then Call DBSet_KZT: Call DBDelete_KZT
    If Not B1Op_HJKbn(0) Then Call DBSet_KJT: Call DBDelete_KJT

    CisDB.DBTran (TransCommit)      ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

    Rv_CallRtn = True
    
    DBDelete = True
    Exit Function
DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False

'  ( 確認メッセージ )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    With CisFun
        .MB_Lines = 5
        If SyoriKB = "ADD" Then
            .MB_MSG(2) = "    追加処理を行います。            "
            .MB_Title = "追加確認"
        Else
            .MB_MSG(2) = "    修正処理を行います。            "
            .MB_Title = "修正確認"
        End If
        .MB_MSG(4) = "     よろしいですか？            "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err

    Call CisDB.DBTran(TransBegin)   ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

' ( 更新処理 )
    If SyoriKB = "ADD" Then
       If Not B1Op_ZGKbn(0) Then Call DBSet_KZT: Call KZTInsert
       If Not B1Op_HJKbn(0) Then Call DBSet_KJT: Call KJTInsert
    Else
       mBefore_ZGKB = 0
       gSL_Select = "Select * From かんばん増減テーブル"
       gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(iH1_Torcd) & "'"
       gSL_Select = gSL_Select & " And   受入 = '" & RTrim(iH1_Ukeir) & "'"
       gSL_Select = gSL_Select & " And   品番 = '" & RTrim(iH1_Hinbn) & "'"
       gSL_Select = gSL_Select & " And   開始日 = '" & RTrim(Old_SKYmd) & "'"
       If KZTRead(gSL_Select, 1) Then
          mBefore_ZGKB = KZT.増減区分
       End If
'          Call DBSet_KZT
'          Call DBUpdate_KZT
'       End If
'       'If Not B1Op_ZGKbn(0) Then Call DBSet_KZT: Call DBUpdate_KZT
'       If Not B1Op_HJKbn(0) Then Call DBSet_KJT: Call DBUpdate_KJT
       gSL_Select = "DELETE FROM かんばん増減テーブル"
       gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(iH1_Torcd) & "'"
       gSL_Select = gSL_Select & " And   受入 = '" & RTrim(iH1_Ukeir) & "'"
       gSL_Select = gSL_Select & " And   品番 = '" & RTrim(iH1_Hinbn) & "'"
       gSL_Select = gSL_Select & " And   開始日 = '" & RTrim(Old_SKYmd) & "'"
       With CisDB
            .ConnectNo = 0
            .SQL = gSL_Select
            If Not .DBExec Then Exit Function
       End With
       
       gSL_Select = "DELETE FROM かんばん平準化テーブル"
       gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(iH1_Torcd) & "'"
       gSL_Select = gSL_Select & " And   受入 = '" & RTrim(iH1_Ukeir) & "'"
       gSL_Select = gSL_Select & " And   品番 = '" & RTrim(iH1_Hinbn) & "'"
       gSL_Select = gSL_Select & " And   開始日 = '" & RTrim(Old_SKYmd) & "'"
       With CisDB
            .ConnectNo = 0
            .SQL = gSL_Select
            If Not .DBExec Then Exit Function
       End With
       
       If Not B1Op_ZGKbn(0) Then Call DBSet_KZT: Call KZTInsert
       If Not B1Op_HJKbn(0) Then Call DBSet_KJT: Call KJTInsert
    End If

    CisDB.DBTran (TransCommit)      ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

    Rv_CallRtn = True
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+------------------------------------+
'+      かんばん増減テーブル・セット
'+------------------------------------+
Private Function DBSet_KZT()
    DBSet_KZT = False
    Call ItemsClearKZT
    With KZT
        .取引先 = RTrim(iH1_Torcd)
        .受入 = RTrim(iH1_Ukeir)
        .品番 = RTrim(iH1_Hinbn)
        .開始日 = RTrim(iB1_SKYmd)
        .終了日 = RTrim(iB1_EKYmd)
        .設定内容区分 = B1Chk_SKbn
        If B1Op_ZGKbn(1) Then .増減区分 = 0
        If B1Op_ZGKbn(2) Then .増減区分 = 1
        .設定枚数 = CisFun.Val2(iB1_SetSu)
        .便当枚数 = CisFun.Val2(iB1_BinSu)
        .作成者 = gTanto
        .更新者 = gTanto
        If .増減区分 = mBefore_ZGKB Then
            .累積枚数 = CisFun.Val2(B1lb_RuiSu)
            If RTrim(B1lb_KanYmd) <> "" Then
               .完了日 = Format(B1lb_KanYmd, "YYYYMMDD")
               If .設定枚数 > .累積枚数 Then
                   .完了日 = ""
               End If
            End If
        Else
           .累積枚数 = 0
           .完了日 = ""
        End If
    End With
    DBSet_KZT = True
End Function
'*--------------------*'
'*     更新処理
'*--------------------*'
Public Function DBUpdate_KZT(Optional DBNo As Integer) As Boolean
    DBUpdate_KZT = False
    With KZT
        gSL_Select = ""
        gSL_Select = gSL_Select & "update かんばん増減テーブル"
        gSL_Select = gSL_Select & "   set 開始日 = '" & .開始日 & "'"
        gSL_Select = gSL_Select & "     , 終了日 = '" & .終了日 & "'"
        gSL_Select = gSL_Select & "     , 設定内容区分 = " & .設定内容区分 & ""
        gSL_Select = gSL_Select & "     , 増減区分 = " & .増減区分 & ""
        gSL_Select = gSL_Select & "     , 設定枚数 = " & .設定枚数 & ""
        gSL_Select = gSL_Select & "     , 便当枚数 = " & .便当枚数 & ""
        If .設定枚数 > .累積枚数 Then
            gSL_Select = gSL_Select & "     , 完了日 = ''"
        End If
        gSL_Select = gSL_Select & "     , 更新日 = getdate()"
        gSL_Select = gSL_Select & "     , 更新者 = '" & RTrim(.更新者) & "'"
        gSL_Select = gSL_Select & "     , 更新端末 = substring(host_name(),1,20)"
        gSL_Select = gSL_Select & " where 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   and 開始日 = '" & RTrim(Old_SKYmd) & "'"
    End With
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    DBUpdate_KZT = True
End Function
'*--------------------*'
'*     削除処理
'*--------------------*'
Public Function DBDelete_KZT(Optional DBNo As Integer) As Boolean
    DBDelete_KZT = False
    With KZT
        gSL_Select = ""
        gSL_Select = gSL_Select & "delete かんばん増減テーブル"
        gSL_Select = gSL_Select & " where 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   and 開始日 = '" & RTrim(.開始日) & "'"
    End With
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    DBDelete_KZT = True
End Function
'+------------------------------------+
'+      かんばん平準化テーブル・セット
'+------------------------------------+
Private Function DBSet_KJT()
    DBSet_KJT = False
    Call ItemsClearKJT
    With KJT
        .取引先 = RTrim(iH1_Torcd)
        .受入 = RTrim(iH1_Ukeir)
        .品番 = RTrim(iH1_Hinbn)
        .開始日 = RTrim(iB1_SKYmd)
        .終了日 = RTrim(iB1_EKYmd)
        .設定内容区分 = B1Chk_SKbn
        If B1Op_HJKbn(1) Then
           .平準化区分 = 2
           .便上限枚数 = CisFun.Val2(iB1_MonSu)
        End If
        If B1Op_HJKbn(2) Then
           .平準化区分 = 1
           .日上限枚数 = CisFun.Val2(iB1_MonSu)
        End If
        If B1Op_HJKbn(3) Then
           .平準化区分 = 3
           .月上限枚数 = CisFun.Val2(iB1_MonSu)
        End If
        .作成者 = gTanto
        .更新者 = gTanto
    End With
    DBSet_KJT = True
End Function
'*--------------------*'
'*     更新処理
'*--------------------*'
Public Function DBUpdate_KJT(Optional DBNo As Integer) As Boolean
    DBUpdate_KJT = False
    With KJT
        gSL_Select = ""
        gSL_Select = gSL_Select & "update かんばん平準化テーブル"
        gSL_Select = gSL_Select & "   set 開始日 = '" & .開始日 & "'"
        gSL_Select = gSL_Select & "     , 終了日 = '" & .終了日 & "'"
        gSL_Select = gSL_Select & "     , 設定内容区分 = " & .設定内容区分 & ""
        gSL_Select = gSL_Select & "     , 平準化区分 = " & .平準化区分 & ""
        gSL_Select = gSL_Select & "     , 便上限枚数 = " & .便上限枚数 & ""
        gSL_Select = gSL_Select & "     , 日上限枚数 = " & .日上限枚数 & ""
        gSL_Select = gSL_Select & "     , 月上限枚数 = " & .月上限枚数 & ""
        gSL_Select = gSL_Select & "     , 更新日 = getdate()"
        gSL_Select = gSL_Select & "     , 更新者 = '" & RTrim(.更新者) & "'"
        gSL_Select = gSL_Select & "     , 更新端末 = substring(host_name(),1,20)"
        gSL_Select = gSL_Select & " where 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   and 開始日 = '" & RTrim(Old_SKYmd) & "'"
    End With
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    DBUpdate_KJT = True
End Function
'*--------------------*'
'*     削除処理
'*--------------------*'
Public Function DBDelete_KJT(Optional DBNo As Integer) As Boolean
    DBDelete_KJT = False
    With KJT
        gSL_Select = ""
        gSL_Select = gSL_Select & "delete かんばん平準化テーブル"
        gSL_Select = gSL_Select & " where 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   and 開始日 = '" & RTrim(.開始日) & "'"
    End With
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    DBDelete_KJT = True
End Function
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理　関　連
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ - かんばん増減管理
'+----------------------------------------------------------+
Private Sub GridInit1()
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 2, 0, "№")
        Call .FixedSet(中中, 中中, あり, 5, 1, "設定")
        Call .FixedSet(中中, 中中, あり, 11, 2, "開 始 日")
        Call .FixedSet(中中, 中中, あり, 11, 3, "終 了 日")
        Call .FixedSet(中中, 中中, あり, 12, 4, "増減")
        Call .FixedSet(中中, 右中, あり, 8, 5, "設定枚数")
        Call .FixedSet(中中, 右中, あり, 8, 6, "便当枚数")
        Call .FixedSet(中中, 右中, あり, 8, 7, "累積枚数")
        Call .FixedSet(中中, 中中, あり, 12, 8, "完 了 日")
        Call .FixedSet(中中, 中中, あり, 10, 9, "平準化")
        Call .FixedSet(中中, 右中, あり, 10, 10, "上限枚数")
        Call .FixedSet(中中, 中中, あり, 11, 11, "更新日")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 12, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11")
        .InitDisp
    End With
End Sub
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ - 内示情報
'+----------------------------------------------------------+
Private Sub GridInit2()
    With CisVsGrid2
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 中中, あり, 8, 1, "内示年月")
        Call .FixedSet(中中, 右中, あり, 7, 2, "内示数")
        Call .FixedSet(中中, 右中, あり, 8, 3, "日当枚数")
        Call .FixedSet(中中, 右中, あり, 8, 4, "回転枚数")
        Call .FixedSet(中中, 右中, あり, 6, 5, "前月差")
        Call .FixedSet(中中, 右中, あり, 8, 6, "平準化日")
        Call .FixedSet(中中, 右中, あり, 8, 7, "平準化便")
        Call .FixedSet(中中, 右中, あり, 0, 8, "稼動日数")
        Call .FixedSet(中中, 右中, あり, 0, 9, "年月")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 10, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7")
        .InitDisp
    End With
End Sub
'****************************************************'
'*****      グリッド内容セット - かんばん増減管理
'****************************************************'
Private Function GridSet1() As Boolean
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

    GridSet1 = False
    Call CisVsGrid1.Clear
    
    With CisDB
         .ConnectNo = 0
         .SQL = "DELETE FROM かんばん増減一括詳細ワーク "
         .SQL = .SQL & " WHERE 処理端末 = '" & RTrim(CisFun.WSID) & "'"
         If Not .DBExec Then Exit Function
    End With
    
    gSL_Select = "INSERT INTO かんばん増減一括詳細ワーク "
    If Trim(iH1_Ym) <> "" Then
       gSL_Select = gSL_Select & " Select "
    Else
       gSL_Select = gSL_Select & " Select Top 5 "
    End If
    gSL_Select = gSL_Select & " '" & RTrim(CisFun.WSID) & "',"
    gSL_Select = gSL_Select & " CASE WHEN KZ.増減区分 = 0 THEN 1 "
    gSL_Select = gSL_Select & "      ELSE 2 END 増減区分,"
    gSL_Select = gSL_Select & " KZ.設定枚数,KZ.便当枚数,KZ.累積枚数,KZ.開始日,KZ.終了日,"
    gSL_Select = gSL_Select & " KZ.設定内容区分 設定内容区分,KZ.完了日,"
    gSL_Select = gSL_Select & " 0 平準化区分,0 平準設定枚数,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(KZ.更新日,'') = '' THEN KZ.作成日 ELSE KZ.更新日 END データ更新日"
'    gSL_Select = gSL_Select & " INTO WK7" & Replace(RTrim(CisFun.WSID), "-", "")
    gSL_Select = gSL_Select & " FROM かんばん増減テーブル KZ"
    gSL_Select = gSL_Select & " WHERE KZ.取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " AND KZ.受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & " AND KZ.品番 = '" & RTrim(iH1_Hinbn) & "'"
    If Trim(iH1_Ym) <> "" Then
       gSL_Select = gSL_Select & " AND SUBSTRING(KZ.開始日, 1, 6) <= '" & iH1_Ym & "'"
       gSL_Select = gSL_Select & " AND CASE WHEN KZ.終了日 = '' THEN '999912' ELSE SUBSTRING(KZ.終了日,1,6) END >= '" & iH1_Ym & "'"
    End If
    gSL_Select = gSL_Select & " UNION "
    If Trim(iH1_Ym) <> "" Then
       gSL_Select = gSL_Select & " Select "
    Else
       gSL_Select = gSL_Select & " Select Top 5 "
    End If
    gSL_Select = gSL_Select & "  '" & RTrim(CisFun.WSID) & "',"
    gSL_Select = gSL_Select & " 0 増減区分,0 設定枚数,0 便当枚数,0 累積枚数,HJ.開始日,HJ.終了日,"
    gSL_Select = gSL_Select & " HJ.設定内容区分 設定内容区分,'' 完了日,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(HJ.品番,'') = '' THEN 0 ELSE HJ.平準化区分 END 平準化区分,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(HJ.平準化区分,0) = 1 THEN HJ.日上限枚数 "
    gSL_Select = gSL_Select & "      WHEN ISNULl(HJ.平準化区分,0) = 2 THEN HJ.便上限枚数 "
    gSL_Select = gSL_Select & "      WHEN ISNULL(HJ.平準化区分,0) = 3 THEN HJ.月上限枚数 "
    gSL_Select = gSL_Select & " END 平準設定枚数,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(HJ.更新日,'') = '' THEN HJ.作成日 ELSE HJ.更新日 END データ更新日"
    gSL_Select = gSL_Select & " FROM かんばん平準化テーブル HJ"
    gSL_Select = gSL_Select & " WHERE HJ.取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " AND HJ.受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & " AND HJ.品番 = '" & RTrim(iH1_Hinbn) & "'"
    If Trim(iH1_Ym) <> "" Then
       gSL_Select = gSL_Select & " AND SUBSTRING(HJ.開始日, 1, 6) <= '" & iH1_Ym & "'"
       gSL_Select = gSL_Select & " AND CASE WHEN HJ.終了日 = '' THEN '999912' ELSE SUBSTRING(HJ.終了日,1,6) END >= '" & iH1_Ym & "'"
    End If
    With CisDB
        .ConnectNo = 0
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    
    gSL_Select = " SELECT "
    gSL_Select = gSL_Select & " MAX(増減区分) 増減区分,MAX(設定枚数) 設定枚数,MAX(便当枚数) 便当枚数,MAX(累積枚数) 累積枚数,開始日,終了日,"
    gSL_Select = gSL_Select & " 設定内容区分,MAX(完了日) 完了日,MAX(平準化区分) 平準化区分,MAX(平準設定枚数) 平準設定枚数,"
    gSL_Select = gSL_Select & " MAX(データ更新日) データ更新日"
'    gSL_Select = gSL_Select & " FROM WK7" & Replace(RTrim(CisFun.WSID), "-", "")
    gSL_Select = gSL_Select & " FROM かんばん増減一括詳細ワーク "
    gSL_Select = gSL_Select & " WHERE 処理端末 = '" & RTrim(CisFun.WSID) & "'"
    gSL_Select = gSL_Select & " GROUP BY"
    gSL_Select = gSL_Select & " 開始日, 終了日, 設定内容区分"
    gSL_Select = gSL_Select & " ORDER BY ISNULL(開始日,'') Desc "
    
    If KRVRead(gSL_Select, , 1) Then

'       抽出処理 初期ｾｯﾄ
        Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
'       ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
        With vsGrid1
           .Redraw = False
            If Ttl_Kensu + 1 < CisVsGrid1.DispMax Then
                .Rows = CisVsGrid1.DispMax
            Else
                .Rows = Ttl_Kensu + 1
            End If
        End With
        ID = 0
        Do Until Not KRV_RDSTS
            ID = ID + 1

            If Rv_Call2 <> "" And mFirst_Flg = 1 Then
               If KRV.開始日 = RV_SYmd And KRV.終了日 = RV_EYmd Then
                  mFirst_Gyo = ID
               End If
            End If
            
            Call KrvInfoToGrid1(ID)

            Call KRVReadNext(1)
        Loop
                
        With vsGrid1
             .Row = 1: .Col = 1: .ColSel = .Cols - 1
             .Redraw = True
        End With

        GridSet1 = True
    End If
    Call KRVClose
End Function
'+----------------------------------------------+
'+  かんばん増減テーブルよりグリッドに転送
'+----------------------------------------------+
Private Sub KrvInfoToGrid1(SetRow As Long)
    With vsGrid1
        'No
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(0)) = SetRow
        '設定
        Select Case KRV.設定内容区分
            Case 0: .TextMatrix(SetRow, CisVsGrid1.FixedGet(1)) = "有効"
            Case 1: .TextMatrix(SetRow, CisVsGrid1.FixedGet(1)) = "無効"
            Case Else: .TextMatrix(SetRow, CisVsGrid1.FixedGet(1)) = ""
        End Select
        '開始日
        If KRV.開始日 <> "" Then
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(2)) = Format(KRV.開始日, "0000/00/00")
        End If
        '終了日
        If KRV.終了日 <> "" Then
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(3)) = Format(KRV.終了日, "0000/00/00")
        End If
        '増減区分
        Select Case KRV.増減区分
            Case 1: .TextMatrix(SetRow, CisVsGrid1.FixedGet(4)) = "減"
            Case 2: .TextMatrix(SetRow, CisVsGrid1.FixedGet(4)) = "増"
            Case Else: .TextMatrix(SetRow, CisVsGrid1.FixedGet(4)) = ""
        End Select
        '設定枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(5)) = Format(KRV.設定枚数, "#,###")
        '便当枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(6)) = Format(KRV.便当枚数, "#,###")
        '累積枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(7)) = Format(KRV.累積枚数, "#,###")
        '完了日
        If KRV.完了日 <> "" Then
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(8)) = Format(KRV.完了日, "0000/00/00")
        End If
        '平準化区分
        Select Case KRV.平準化区分
            Case 1: .TextMatrix(SetRow, CisVsGrid1.FixedGet(9)) = "日"
            Case 2: .TextMatrix(SetRow, CisVsGrid1.FixedGet(9)) = "便"
            Case 3: .TextMatrix(SetRow, CisVsGrid1.FixedGet(9)) = "月"
            Case Else: .TextMatrix(SetRow, CisVsGrid1.FixedGet(9)) = ""
        End Select
        '平準上限枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(10)) = Format(KRV.平準設定枚数, "#,###")
        '更新日
        If Format(KRV.データ更新日, "0000/00/00") <> "" Then
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(11)) = Format(KRV.データ更新日, "YYYY/MM/DD")
        End If
    End With
End Sub
'******************************************'
'*****      行選択（クリック）
'******************************************'
Private Sub vsGrid1_Click()
    If SyoriKB <> "ADD" Then
        mMntRow = vsGrid1.Row
        Call FromGrid1
        Exit Sub
    Else
        SendKeys "{TAB}"
    End If
End Sub
'******************************************'
'*****      行選択（スペース）
'******************************************'
Private Sub vsGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If SyoriKB <> "ADD" Then
        If KeyCode = vbKeySpace Then
            mMntRow = vsGrid1.Row
            Call FromGrid1
            Exit Sub
        End If
    Else
        SendKeys "{TAB}"
    End If
End Sub
'+-------------------------------+
'+      行変更時　内容表示
'+-------------------------------+
Private Sub VsGrid1_RowColChange()
    If SyoriKB <> "ADD" Then
        Call BodySet1
    End If
End Sub
'******************************************'
'*****      行選択（クリック） - 内示情報
'******************************************'
Private Sub vsGrid2_Click()
    SendKeys "{TAB}"
End Sub
'******************************************'
'*****      平準化数　再計算
'******************************************'
Private Function NewHeiSu() As Integer
    Dim wBinSu          As Byte
    Dim wHeiKeisu       As Integer
    Dim wKadoDay        As Double
    Dim wKaitenSu       As Integer
    Dim wDaySu          As Double               '2006/03/20
    Dim wC1             As Integer              '2006/03/20
    Dim wC2             As Integer              '2006/03/20
    Dim wC3             As Integer              '2006/03/20
    Dim wCKeisu         As Double               '2006/03/20
    Dim wKKeisu         As Double               '2006/03/20
    
    NewHeiSu = 0
    
    If B1Op_HJKbn(0) Then Exit Function
    
    wBinSu = CisFun.Val2(Mid(B1lb_Cycle, 3, 2))                 '便数
    wHeiKeisu = CisFun.Val2(B1lb_HKeisu)                        '平準化係数
    
'-------------------------------------------- 2006/03/22 Delete
'    gSL_Select = "SELECT * FROM カレンダマスタ "
'    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
'    gSL_Select = gSL_Select & " AND 取引先 = ''"
'    gSL_Select = gSL_Select & " AND 受入 = ''"
'    gSL_Select = gSL_Select & " AND 年 = '" & RTrim(iB1_SKYmd.cYear) & "'"
'    gSL_Select = gSL_Select & " AND 月 = '" & RTrim(iB1_SKYmd.cMonth) & "'"
'    Call CLMRead(gSL_Select, 1)
'    wKadoDay = CLM.稼動日数                                     '稼働日数
'-------------------------------------------- 2006/03/22 Delete End
    
'-------------------------------------------- 2006/03/16 Delete
'    wKaitenSu = CisFun.Val2(B1lb_KaiMai)
'-------------------------------------------- 2006/03/16 Delete End
    wKaitenSu = 0
    If CisFun.Val2(B1lb_YSu(0)) <> 0 Then
       wKaitenSu = CisFun.Val2(B1lb_YSu(0))
    End If
'-------------------------------------------- 2006/03/16 Delete
'    If CisFun.Val2(B1lb_YSu(1)) <> 0 Then
'       wKaitenSu = CisFun.Val2(B1lb_YSu(1))
'    End If
'-------------------------------------------- 2006/03/16 Delete End
'
'-------------------------------------------- 2006/03/20 Delete
'    If B1Op_ZGKbn(1) Then
'       wKaitenSu = wKaitenSu - (CisFun.Val2(iB1_SetSu) - CisFun.Val2(B1lb_RuiSu))
'    Else
'       If B1Op_ZGKbn(2) Then
'          wKaitenSu = wKaitenSu + (CisFun.Val2(iB1_SetSu) - CisFun.Val2(B1lb_RuiSu))
'       End If
'    End If
'
'    '便上限枚数
'    If B1Op_HJKbn(1).Value Then
''       CEILING( WK.回転枚数 / CAST( WK.便数 AS DECIMAL( 3,1 ) ) * ( 1 + WK.平準化係数 /100.0 ) )
'       NewHeiSu = CisFun.HasuProc(wKaitenSu / wBinSu * (1 + wHeiKeisu / 100), 切上げ)
'    End If
'    '日上限枚数
'    If B1Op_HJKbn(2).Value Then
''       CEILING( WK.回転枚数 * ( 1 + WK.平準化係数 /100.0 ) )
'       NewHeiSu = CisFun.HasuProc(wKaitenSu * (1 + wHeiKeisu / 100), 切上げ)
'    End If
'    '月上限枚数
'    If B1Op_HJKbn(3).Value Then
''       CEILING( WK.回転枚数 * WK.稼動日数 * ( 1 + WK.平準化係数 /100.0 ) )
'       NewHeiSu = CisFun.HasuProc(wKaitenSu * wKadoDay * (1 + wHeiKeisu / 100), 切上げ)
'    End If
'-------------------------------------------- 2006/03/20 Delete End
'-------------------------------------------- 2006/03/20 Add
'-------------------------------------------- 2006/03/21 Insert
    If B1Op_ZGKbn(1) Then
       wKaitenSu = wKaitenSu - (CisFun.Val2(iB1_SetSu) - CisFun.Val2(B1lb_RuiSu))
    Else
       If B1Op_ZGKbn(2) Then
          wKaitenSu = wKaitenSu + (CisFun.Val2(iB1_SetSu) - CisFun.Val2(B1lb_RuiSu))
       End If
    End If
'    If B1Op_ZGKbn(1) Then
'       wKaitenSu = wKaitenSu - CisFun.Val2(iB1_SetSu)
'    Else
'       If B1Op_ZGKbn(2) Then
'          wKaitenSu = wKaitenSu + CisFun.Val2(iB1_SetSu)
'       End If
'    End If
'-------------------------------------------- 2006/03/21 Insert End
    'サイクル係数算出
    wDaySu = 0
    wKKeisu = CisFun.Val2(B1lb_KKeisu)
    wC1 = CisFun.Val2(Mid(B1lb_Cycle, 1, 1))
    wC2 = CisFun.Val2(Mid(B1lb_Cycle, 3, 2))
    wC3 = CisFun.Val2(Mid(B1lb_Cycle, 6, 2))
    wCKeisu = ((wC1 * (wC3 + 1)) / wC2) + wKKeisu
    wDaySu = wKaitenSu / wCKeisu
    wDaySu = CisFun.HasuProc(wDaySu, 切上げ)
    '便上限枚数
    If B1Op_HJKbn(1).Value Then
       NewHeiSu = CisFun.HasuProc(wDaySu / wBinSu * (1 + wHeiKeisu / 100), 切上げ)
    End If
    '日上限枚数
    If B1Op_HJKbn(2).Value Then
       NewHeiSu = CisFun.HasuProc(wDaySu * (1 + wHeiKeisu / 100), 切上げ)
    End If
    '月上限枚数
    If B1Op_HJKbn(3).Value Then
'-------------------------------------------- 2006/03/22 Insert
       gSL_Select = "SELECT * FROM カレンダマスタ "
       gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
       gSL_Select = gSL_Select & " AND 取引先 = ''"
       gSL_Select = gSL_Select & " AND 受入 = ''"
       gSL_Select = gSL_Select & " AND 年 = '" & RTrim(iB1_SKYmd.cYear) & "'"
       gSL_Select = gSL_Select & " AND 月 = '" & RTrim(iB1_SKYmd.cMonth) & "'"
       Call CLMRead(gSL_Select, 1)
       wKadoDay = CLM.稼動日数                                     '稼働日数
'-------------------------------------------- 2006/03/22 Insert End
       NewHeiSu = CisFun.HasuProc(wDaySu * wKadoDay * (1 + wHeiKeisu / 100), 切上げ)
    End If
'-------------------------------------------- 2006/03/20 Add End
    
    If NewHeiSu < 0 Then NewHeiSu = 0
    
    '平準枚数差
'''    NewHeiSu = NewHeiSu - CisFun.Val2(iB1_MonSu)
End Function
'***************************************************************************************
'* グリッド内容をワーク項目にセット
'***************************************************************************************
Private Sub Grid_Item_Set()
    Dim ID          As Integer
    Dim Ix          As Integer
    
    iH1_Ym = RV_YM
    Call BodySet(All)
    
    With vsGrid1
        RVS_ZG_Count = 0
        For ID = 1 To .Rows - 1
            If .TextMatrix(ID, CisVsGrid1.FixedGet(0)) = "" Then Exit For
            If .TextMatrix(ID, CisVsGrid1.FixedGet(2)) = "" Then Exit For
'            RVS_ZG_Count = RVS_ZG_Count + 1
            If RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(1))) = "有効" Then         '【有効】のみ対象 2006/03/20
               RVS_ZG_Count = RVS_ZG_Count + 1
            End If
        Next ID
        
        If RVS_ZG_Count <> 0 Then
            ReDim RVS_ZG_SYMD(RVS_ZG_Count)
            ReDim RVS_ZG_EYMD(RVS_ZG_Count)
            ReDim RVS_ZG_ZGKB(RVS_ZG_Count)
            ReDim RVS_ZG_ZGNKB(RVS_ZG_Count)
            ReDim RVS_ZG_ZGSet(RVS_ZG_Count)
            ReDim RVS_ZG_ZGBin(RVS_ZG_Count)
            ReDim RVS_ZG_HJKB(RVS_ZG_Count)
            ReDim RVS_ZG_HJNKB(RVS_ZG_Count)
            ReDim RVS_ZG_HJSet(RVS_ZG_Count)
            ReDim RVS_ZG_Kan(RVS_ZG_Count)
            ReDim RVS_ZG_RuiMai(RVS_ZG_Count)
            
            Ix = 0
            
            For ID = 1 To .Rows - 1
                If .TextMatrix(ID, CisVsGrid1.FixedGet(0)) = "" Then Exit For
                If .TextMatrix(ID, CisVsGrid1.FixedGet(2)) = "" Then Exit For
                
                If RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(1))) = "有効" Then         '【有効】のみ対象 2006/03/20
                    
                    Select Case RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(1)))
                        Case "有効": RVS_ZG_ZGNKB(Ix) = 0: RVS_ZG_ZGNKB(Ix) = 0
                        Case "無効": RVS_ZG_ZGNKB(Ix) = 1: RVS_ZG_ZGNKB(Ix) = 1
                    End Select
                    RVS_ZG_SYMD(Ix) = Format(.TextMatrix(ID, CisVsGrid1.FixedGet(2)), "YYYYMMDD")
                    RVS_ZG_EYMD(Ix) = Format(.TextMatrix(ID, CisVsGrid1.FixedGet(3)), "YYYYMMDD")
                    Select Case RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(4)))
                        Case "減": RVS_ZG_ZGKB(Ix) = 1
                        Case "増": RVS_ZG_ZGKB(Ix) = 2
                        Case Else: RVS_ZG_ZGKB(Ix) = 0
                    End Select
                    RVS_ZG_ZGSet(Ix) = CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(5)))
                    RVS_ZG_ZGBin(Ix) = CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(6)))
                    Select Case RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(9)))
                        Case "日": RVS_ZG_HJKB(Ix) = 1
                        Case "便": RVS_ZG_HJKB(Ix) = 2
                        Case "月": RVS_ZG_HJKB(Ix) = 3
                        Case Else: RVS_ZG_HJKB(Ix) = 0
                    End Select
                    RVS_ZG_HJSet(Ix) = CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(10)))
                    If RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(8))) <> "" Then
                       RVS_ZG_Kan(Ix) = Format(.TextMatrix(ID, CisVsGrid1.FixedGet(8)), "YYYYMMDD")
                    Else
                       RVS_ZG_Kan(Ix) = ""
                    End If
                    RVS_ZG_RuiMai(Ix) = CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(7)))
                
                    Ix = Ix + 1
                    
                Else
                
                    RVS_ZG_Count = RVS_ZG_Count - 1
                
                End If
                
            Next ID
        Else
            RVS_ZG_Count = 1
            
            ReDim RVS_ZG_SYMD(RVS_ZG_Count)
            ReDim RVS_ZG_EYMD(RVS_ZG_Count)
            ReDim RVS_ZG_ZGKB(RVS_ZG_Count)
            ReDim RVS_ZG_ZGNKB(RVS_ZG_Count)
            ReDim RVS_ZG_ZGSet(RVS_ZG_Count)
            ReDim RVS_ZG_ZGBin(RVS_ZG_Count)
            ReDim RVS_ZG_HJKB(RVS_ZG_Count)
            ReDim RVS_ZG_HJNKB(RVS_ZG_Count)
            ReDim RVS_ZG_HJSet(RVS_ZG_Count)
            ReDim RVS_ZG_Kan(RVS_ZG_Count)
            ReDim RVS_ZG_RuiMai(RVS_ZG_Count)
            
            ID = 1
            
' --------------------- 2006/03/20 Delete
'
'                Select Case RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(1)))
'                    Case "有効": RVS_ZG_ZGNKB(ID - 1) = 0: RVS_ZG_ZGNKB(ID - 1) = 0
'                    Case "無効": RVS_ZG_ZGNKB(ID - 1) = 1: RVS_ZG_ZGNKB(ID - 1) = 1
'                    Case Else: RVS_ZG_ZGNKB(ID - 1) = 0: RVS_ZG_ZGNKB(ID - 1) = 0
'                End Select
''                RVS_ZG_SYMD(ID - 1) = Format(.TextMatrix(ID, CisVsGrid1.FixedGet(2)), "YYYYMMDD")
''                RVS_ZG_EYMD(ID - 1) = Format(.TextMatrix(ID, CisVsGrid1.FixedGet(3)), "YYYYMMDD")
'                RVS_ZG_SYMD(ID - 1) = "2"
'                RVS_ZG_EYMD(ID - 1) = ""
'                Select Case RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(4)))
'                    Case "減": RVS_ZG_ZGKB(ID - 1) = 1
'                    Case "増": RVS_ZG_ZGKB(ID - 1) = 2
'                    Case Else: RVS_ZG_ZGKB(ID - 1) = 0
'                End Select
'                RVS_ZG_ZGSet(ID - 1) = CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(5)))
'                RVS_ZG_ZGBin(ID - 1) = CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(6)))
'                Select Case RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(9)))
'                    Case "日": RVS_ZG_HJKB(ID - 1) = 1
'                    Case "便": RVS_ZG_HJKB(ID - 1) = 2
'                    Case "月": RVS_ZG_HJKB(ID - 1) = 3
'                    Case Else: RVS_ZG_HJKB(ID - 1) = 0
'                End Select
'                RVS_ZG_HJSet(ID - 1) = CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(10)))
'                If RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(8))) <> "" Then
'                   RVS_ZG_Kan(ID - 1) = Format(.TextMatrix(ID, CisVsGrid1.FixedGet(8)), "YYYYMMDD")
'                Else
'                   RVS_ZG_Kan(ID - 1) = ""
'                End If
'                RVS_ZG_RuiMai(ID - 1) = CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(7)))
'
' --------------------- 2006/03/20 Delete End
                
        End If
    End With
    
End Sub
