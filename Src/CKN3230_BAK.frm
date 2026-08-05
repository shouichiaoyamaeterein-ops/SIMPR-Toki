VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CKN3230 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "工程かんばん個別発行"
   ClientHeight    =   10905
   ClientLeft      =   90
   ClientTop       =   2010
   ClientWidth     =   15240
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
   ScaleHeight     =   10905
   ScaleWidth      =   15240
   Begin VB.PictureBox B3_Area1 
      BackColor       =   &H00FF0000&
      Height          =   1965
      Left            =   6585
      ScaleHeight     =   1905
      ScaleWidth      =   3645
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4830
      Visible         =   0   'False
      Width           =   3705
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   1665
         Index           =   1
         Left            =   120
         Top             =   120
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   2937
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
         Begin Cis3D_v60.CIS3D CIS3D21 
            Height          =   405
            Index           =   1
            Left            =   15
            Top             =   0
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   714
            BackColor       =   16761024
            Caption         =   "【　テスト印刷　】"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D22 
            Height          =   540
            Left            =   0
            Top             =   420
            Width           =   3420
            _ExtentX        =   6033
            _ExtentY        =   953
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
            Begin Cis3D_v60.CIS3D CIS3D29 
               Height          =   480
               Left            =   2730
               Top             =   15
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   847
               ForeColor       =   16711680
               Caption         =   "枚"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
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
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   435
               Index           =   0
               Left            =   615
               Top             =   30
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   767
               ForeColor       =   16711680
               Caption         =   "発行枚数"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   12632256
               cBoderColor2    =   12632256
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cBoderStyle     =   2
               cPositionY      =   30
            End
            Begin CisText_V60.CisText iB3_HSu 
               Height          =   360
               Left            =   2130
               TabIndex        =   18
               Top             =   90
               Width           =   525
               _ExtentX        =   926
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
               Text            =   "  3"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cFdAutoFormat   =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D PB_Can2 
            Height          =   705
            Left            =   15
            Top             =   960
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   1244
            ForeColor       =   49152
            Caption         =   " 取消 (ESC)"
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
            cBoderWidth     =   4
            cAlingnment     =   6
         End
         Begin Cis3D_v60.CIS3D PB_ENT2 
            Height          =   705
            Left            =   1725
            Top             =   960
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   1244
            ForeColor       =   192
            Caption         =   " 開始 (Ent)"
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
            cBoderWidth     =   4
            cAlingnment     =   6
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'なし
      Height          =   495
      Index           =   1
      Left            =   11220
      ScaleHeight     =   495
      ScaleWidth      =   2520
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6330
      Width           =   2520
      Begin Cis3D_v60.CIS3D UC_3D6 
         Height          =   435
         Left            =   30
         Top             =   30
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   767
         ForeColor       =   0
         Caption         =   " 発行総枚数"
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
         cAlingnment     =   1
         Begin Cis3D_v60.CIS3D iT1_Addsu 
            Height          =   315
            Left            =   1350
            Top             =   60
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BackColor       =   16761024
            ForeColor       =   0
            Caption         =   "ZZ,ZZ9"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
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
   End
   Begin Cis3D_v60.CIS3D B1_Area2 
      Height          =   2235
      Left            =   1500
      Top             =   2340
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   3942
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
      Begin Cis3D_v60.CIS3D CIS3D23 
         Height          =   705
         Left            =   5055
         Top             =   765
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1244
         Caption         =   "部品区分"
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
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_BuKB 
            Height          =   330
            Left            =   240
            Top             =   315
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "完成品"
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
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   705
         Left            =   45
         Top             =   1470
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   1244
         Caption         =   "得意先"
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
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_Tokui 
            Height          =   360
            Left            =   60
            Top             =   315
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "テクノエイト"
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
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   705
         Left            =   6255
         Top             =   1470
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1244
         Caption         =   "ﾎﾟｲﾝﾄ"
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
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D B1lb_Point 
            Height          =   330
            Left            =   150
            Top             =   315
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "01"
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
         Begin Cis3D_v60.CIS3D NoImage_Point 
            Height          =   495
            Left            =   810
            Top             =   120
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   873
            BackColor       =   14737632
            Caption         =   " No  Image"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "HGSｺﾞｼｯｸE"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   14737632
            cBoderColor2    =   14737632
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   2
            cPositionY      =   -70
         End
         Begin VB.Image Image_Point 
            BorderStyle     =   1  '実線
            Height          =   555
            Left            =   780
            Picture         =   "CKN3230.frx":0000
            Stretch         =   -1  'True
            Top             =   90
            Width           =   690
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   705
         Left            =   2910
         Top             =   60
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   1244
         Caption         =   "⇔"
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
         cAlingnment     =   7
         cPositionY      =   100
      End
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   705
         Left            =   6870
         Top             =   60
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1244
         ForeColor       =   0
         Caption         =   "車型"
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
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_Syakei 
            Height          =   330
            Left            =   135
            Top             =   315
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "117W"
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
      Begin Cis3D_v60.CIS3D CIS3D12 
         Height          =   705
         Left            =   45
         Top             =   765
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   1244
         ForeColor       =   0
         Caption         =   "品    名"
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
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_Hinnm 
            Height          =   330
            Left            =   60
            Top             =   315
            Width           =   4890
            _ExtentX        =   8625
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "117;ﾘﾔｻｽ"
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
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   705
         Index           =   0
         Left            =   45
         Top             =   60
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   1244
         ForeColor       =   0
         Caption         =   "返却場所"
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
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_Henkyaku 
            Height          =   330
            Left            =   135
            Top             =   330
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   582
            BackColor       =   12648384
            Caption         =   "A-124"
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
      Begin Cis3D_v60.CIS3D CIS3D15 
         Height          =   705
         Index           =   0
         Left            =   2265
         Top             =   1470
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1244
         ForeColor       =   0
         Caption         =   "生産ﾛｯﾄ"
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
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_SLot 
            Height          =   330
            Index           =   0
            Left            =   285
            Top             =   315
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "1,000"
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
            cPositionX      =   -30
         End
         Begin Cis3D_v60.CIS3D B1lb_SLot 
            Height          =   330
            Index           =   1
            Left            =   1275
            Top             =   330
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
            BackColor       =   16777152
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
         Begin Cis3D_v60.CIS3D CIS3D8 
            Height          =   360
            Left            =   135
            Top             =   315
            Width           =   150
            _ExtentX        =   265
            _ExtentY        =   635
            BackColor       =   12648384
            Caption         =   "1"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D11 
            Height          =   330
            Left            =   1125
            Top             =   330
            Width           =   150
            _ExtentX        =   265
            _ExtentY        =   582
            BackColor       =   12648384
            Caption         =   "2"
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
            cAlingnment     =   7
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D24 
         Height          =   705
         Left            =   6870
         Top             =   765
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1244
         ForeColor       =   0
         Caption         =   "生産区分"
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
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_SeiKu 
            Height          =   330
            Left            =   330
            Top             =   315
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "B"
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
      Begin Cis3D_v60.CIS3D CIS3D34 
         Height          =   705
         Left            =   4560
         Top             =   1470
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1244
         ForeColor       =   0
         Caption         =   "所番地"
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
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_Okiba 
            Height          =   330
            Left            =   165
            Top             =   315
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "1S-A-1"
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
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   705
         Left            =   3270
         Top             =   60
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   1244
         ForeColor       =   0
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D CIS3D14 
            Height          =   360
            Index           =   0
            Left            =   105
            Top             =   210
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   635
            Caption         =   "後工程"
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
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D B1lb_AKotei 
            Height          =   300
            Index           =   0
            Left            =   915
            Top             =   30
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   529
            BackColor       =   16777152
            Caption         =   "スポット"
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
         Begin Cis3D_v60.CIS3D B1lb_AKotei 
            Height          =   300
            Index           =   1
            Left            =   915
            Top             =   360
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   529
            BackColor       =   16777152
            Caption         =   "S-11"
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
      Begin Cis3D_v60.CIS3D NoImage 
         Height          =   2025
         Left            =   7965
         Top             =   105
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   3572
         BackColor       =   14737632
         Caption         =   "No Image"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "HGSｺﾞｼｯｸE"
            Size            =   15.75
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
         cAlingnment     =   7
         cBoderStyle     =   2
      End
      Begin VB.Image Image_Hinban 
         BorderStyle     =   1  '実線
         Height          =   2115
         Left            =   7905
         Stretch         =   -1  'True
         Top             =   60
         Width           =   3000
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
      Height          =   945
      Left            =   1500
      ScaleHeight     =   885
      ScaleWidth      =   9060
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1305
      Width           =   9120
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   780
         Left            =   60
         Top             =   60
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1376
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
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Seban 
            Height          =   360
            Left            =   105
            TabIndex        =   0
            Top             =   345
            Width           =   1095
            _ExtentX        =   1931
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
         Height          =   780
         Left            =   5655
         Top             =   60
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   1376
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
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   360
            Left            =   90
            TabIndex        =   3
            Top             =   330
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
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   780
         Index           =   1
         Left            =   1365
         Top             =   60
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "社内手配先"
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
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   60
            TabIndex        =   1
            Top             =   330
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
            Left            =   1035
            Top             =   315
            Width           =   2475
            _ExtentX        =   4366
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
         Height          =   780
         Index           =   1
         Left            =   4950
         Top             =   60
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   1376
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
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   360
            Left            =   150
            TabIndex        =   2
            Top             =   330
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
      Height          =   750
      Left            =   5940
      ScaleHeight     =   690
      ScaleWidth      =   8565
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   8625
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   6000
         Top             =   30
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
         Left            =   6840
         Top             =   30
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
         Left            =   7680
         Top             =   30
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
         Left            =   3495
         Top             =   30
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
         Left            =   2655
         Top             =   30
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
         Left            =   1800
         Top             =   30
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
      Begin CisBtn_60.CisBtn PB_Prt 
         Height          =   585
         Left            =   4335
         Top             =   30
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
      Begin CisBtn_60.CisBtn PB_Del 
         Height          =   585
         Left            =   915
         Top             =   30
         Visible         =   0   'False
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
      Begin CisBtn_60.CisBtn PB_Add 
         Height          =   585
         Left            =   45
         Top             =   30
         Visible         =   0   'False
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
         BtnCaption      =   "追加"
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
      Begin CisBtn_60.CisBtn PB_Test 
         Height          =   585
         Left            =   5175
         Top             =   30
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
         BtnCaption      =   "ﾃｽﾄ"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   0
         BtnShadowColor  =   16777215
         PFCaption       =   "( F9 )"
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
         PFBackColor     =   0
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
      Caption         =   "【 工程かんばん発行  】"
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
         Left            =   14310
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
         Height          =   285
         Left            =   11805
         Top             =   60
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   503
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
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   65535
         Caption         =   "追加"
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
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   1605
      Left            =   1500
      ScaleHeight     =   1545
      ScaleWidth      =   10935
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4665
      Width           =   10995
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   1425
         Left            =   60
         Top             =   60
         Width           =   10815
         _ExtentX        =   19076
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderStyle     =   1
         Begin Cis3D_v60.CIS3D CIS3D28 
            Height          =   1050
            Left            =   15
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1852
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   60
            Begin Cis3D_v60.CIS3D B1lb_SHYmd 
               Height          =   420
               Left            =   195
               Top             =   300
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   741
               BackColor       =   16777152
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
            Begin Cis3D_v60.CIS3D B1lb_SHBin 
               Height          =   420
               Left            =   1605
               Top             =   300
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   741
               BackColor       =   16777152
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
         End
         Begin Cis3D_v60.CIS3D BLB_Sur 
            Height          =   525
            Index           =   1
            Left            =   2190
            Top             =   360
            Width           =   7365
            _ExtentX        =   12991
            _ExtentY        =   926
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   510
               Index           =   5
               Left            =   2100
               Top             =   0
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   900
               ForeColor       =   0
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                  Height          =   330
                  Index           =   0
                  Left            =   210
                  Top             =   75
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "50"
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
                  cPositionX      =   -100
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   510
               Index           =   6
               Left            =   255
               Top             =   0
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   900
               ForeColor       =   0
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin Cis3D_v60.CIS3D B1lb_Syuki 
                  Height          =   330
                  Index           =   0
                  Left            =   210
                  Top             =   75
                  Width           =   1290
                  _ExtentX        =   2275
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "TP331"
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
                  cPositionX      =   30
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   510
               Index           =   7
               Left            =   3660
               Top             =   0
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   900
               ForeColor       =   16711680
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin CisText_V60.CisText iB1_HSu 
                  Height          =   360
                  Index           =   0
                  Left            =   585
                  TabIndex        =   11
                  Top             =   75
                  Width           =   525
                  _ExtentX        =   926
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
                  Text            =   "  3"
                  MaxLength       =   3
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
            End
            Begin Cis3D_v60.CIS3D Back_SSu 
               Height          =   510
               Index           =   0
               Left            =   5325
               Top             =   0
               Width           =   2040
               _ExtentX        =   3598
               _ExtentY        =   900
               ForeColor       =   16711680
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin CisText_V60.CisText iB1_SSSu 
                  Height          =   360
                  Index           =   0
                  Left            =   315
                  TabIndex        =   12
                  Top             =   75
                  Width           =   540
                  _ExtentX        =   953
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
                  Text            =   " "
                  MaxLength       =   3
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
               Begin CisText_V60.CisText iB1_ESSu 
                  Height          =   360
                  Index           =   0
                  Left            =   1200
                  TabIndex        =   13
                  Top             =   75
                  Width           =   540
                  _ExtentX        =   953
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
                  Text            =   " "
                  MaxLength       =   3
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
               Begin Cis3D_v60.CIS3D CIS3D17 
                  Height          =   255
                  Left            =   945
                  Top             =   120
                  Width           =   210
                  _ExtentX        =   370
                  _ExtentY        =   450
                  Caption         =   "～"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   9
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
                  cBoderStyle     =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   510
               Left            =   15
               Top             =   0
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   900
               BackColor       =   12648384
               Caption         =   "1"
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
               cAlingnment     =   7
            End
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   1395
            Index           =   1
            Left            =   9555
            Top             =   15
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   2461
            ForeColor       =   0
            Caption         =   "最終№"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
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
            Begin Cis3D_v60.CIS3D B1lb_EndSu 
               Height          =   615
               Left            =   120
               Top             =   510
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   1085
               BackColor       =   16777152
               Caption         =   "15"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   18
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
         Begin Cis3D_v60.CIS3D CIS3D19 
            Height          =   345
            Left            =   15
            Top             =   15
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   609
            Caption         =   "収容変更日･便"
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
            cPositionY      =   30
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   345
            Index           =   2
            Left            =   4290
            Top             =   15
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   609
            ForeColor       =   0
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   30
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   345
            Index           =   50
            Left            =   2190
            Top             =   15
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   609
            ForeColor       =   0
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   30
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   345
            Index           =   4
            Left            =   7515
            Top             =   15
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   609
            ForeColor       =   16711680
            Caption         =   "再発行№"
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
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   345
            Index           =   3
            Left            =   5850
            Top             =   15
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   609
            ForeColor       =   16711680
            Caption         =   "追加枚数"
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
         End
         Begin Cis3D_v60.CIS3D BLB_Sur 
            Height          =   525
            Index           =   0
            Left            =   2190
            Top             =   885
            Width           =   7365
            _ExtentX        =   12991
            _ExtentY        =   926
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   510
               Index           =   10
               Left            =   2100
               Top             =   15
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   900
               ForeColor       =   0
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                  Height          =   330
                  Index           =   1
                  Left            =   210
                  Top             =   75
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   " "
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
                  cPositionX      =   -100
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   510
               Index           =   11
               Left            =   255
               Top             =   15
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   900
               ForeColor       =   0
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin Cis3D_v60.CIS3D B1lb_Syuki 
                  Height          =   330
                  Index           =   1
                  Left            =   225
                  Top             =   75
                  Width           =   1290
                  _ExtentX        =   2275
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   " "
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
                  cPositionX      =   30
               End
            End
            Begin Cis3D_v60.CIS3D Back_HSu 
               Height          =   510
               Index           =   1
               Left            =   3660
               Top             =   15
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   900
               ForeColor       =   16711680
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin CisText_V60.CisText iB1_HSu 
                  Height          =   360
                  Index           =   1
                  Left            =   585
                  TabIndex        =   14
                  Top             =   75
                  Width           =   525
                  _ExtentX        =   926
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
                  Text            =   " "
                  MaxLength       =   3
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
            End
            Begin Cis3D_v60.CIS3D Back_SSu 
               Height          =   510
               Index           =   1
               Left            =   5325
               Top             =   15
               Width           =   2040
               _ExtentX        =   3598
               _ExtentY        =   900
               ForeColor       =   16711680
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin CisText_V60.CisText iB1_SSSu 
                  Height          =   360
                  Index           =   1
                  Left            =   330
                  TabIndex        =   15
                  Top             =   75
                  Width           =   540
                  _ExtentX        =   953
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
                  Text            =   " "
                  MaxLength       =   3
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
               Begin CisText_V60.CisText iB1_ESSu 
                  Height          =   360
                  Index           =   1
                  Left            =   1215
                  TabIndex        =   16
                  Top             =   75
                  Width           =   540
                  _ExtentX        =   953
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
                  Text            =   " "
                  MaxLength       =   3
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
               Begin Cis3D_v60.CIS3D CIS3D26 
                  Height          =   255
                  Left            =   960
                  Top             =   120
                  Width           =   210
                  _ExtentX        =   370
                  _ExtentY        =   450
                  Caption         =   "～"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   9
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
                  cBoderStyle     =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D27 
               Height          =   510
               Left            =   15
               Top             =   15
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   900
               BackColor       =   12648384
               Caption         =   "2"
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
               cAlingnment     =   7
            End
         End
      End
   End
   Begin Cis3D_v60.CIS3D B2_Area1 
      Height          =   3840
      Left            =   1500
      Top             =   6840
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   6773
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
      Begin Cis3D_v60.CIS3D CIS3D35 
         Height          =   3660
         Left            =   90
         Top             =   75
         Width           =   12030
         _ExtentX        =   21220
         _ExtentY        =   6456
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
         Begin vsFlexLib.vsFlexArray vsGrid1 
            Height          =   3570
            Left            =   45
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   45
            Width           =   11925
            _Version        =   196608
            _ExtentX        =   21034
            _ExtentY        =   6297
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
            FormatString    =   "№ |手配先 |手配先名(社内)　　|受入|品　番　　　　　　　　　|背番号|収容器   |収容数  |発行枚数||"
            Rows            =   11
            Cols            =   11
            BackColor       =   12648447
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   0
            RowHeightMin    =   320
            GridLines       =   2
            SelectionMode   =   1
         End
      End
   End
   Begin VB.Label KanbanColor2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "購入品"
      Height          =   225
      Left            =   13680
      TabIndex        =   21
      Top             =   2550
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label KanbanColor1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      Caption         =   "仕掛品"
      Height          =   225
      Left            =   13695
      TabIndex        =   20
      Top             =   2265
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label KanbanColor0 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "完成品"
      Height          =   225
      Left            =   13710
      TabIndex        =   10
      Top             =   1980
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label HenColor 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      Caption         =   "対象行色"
      Height          =   225
      Left            =   930
      TabIndex        =   9
      Top             =   6300
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "CKN3230"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   工程かんばん発行
'**       フォームID    :   CKN3230
'**       処理概要      :
'**
'**       作  成  日    :   2009/10/20  By CIS
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)
    
    Dim mWakuColor0     As Long
    Dim mWakuColor1     As Long
    Dim mWakuColor2     As Long
    
    Dim CisVsGrid       As New CisVsGrid3
    Dim CisVsGrid2      As New CisVsGrid3
    
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
    
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 9
        .InitGet = False
        .Init
    End With
    Call GridInit
    
    '+-----------------------+
    '+     色設定 & 獲得     +
    '+-----------------------+
    Call ColorGet
    
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
'
    FormAct = Not FormAct
' イメージ格納先
    With CisFun
'-----------------------------------------------< 2005/12/13 ADD START >
        .INI_FileID = gIniExe & gwsIniName
'-----------------------------------------------< 2005/12/13 ADD START >
        .INI_Section = "Option"
        .INI_Key = "Image":       .INI_Default = gIniExe
        .GetIni
        If Right(RTrim(.INI_String), 1) <> "\" Then
            CKN3230_ImagePath = RTrim(.INI_String) & "\"
        Else
            CKN3230_ImagePath = RTrim(.INI_String)
        End If
'-----------------------------------------------< 2005/12/13 ADD START >
        .INI_FileID = gIniExe & gIniName
'-----------------------------------------------< 2005/12/13 ADD START >
    End With
'
    Call GridSet
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
           Case vbKeyF1:       If PB_Add.Visible Then Call PB_Add_Click    '【新規】
           Case vbKeyF3:       If PB_Del.Visible Then Call PB_Del_Click    '【削除】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyF8:       If PB_Prt.Visible Then Call PB_Prt_Click    '【印刷】
           Case vbKeyF9:       If PB_Test.Visible Then Call PB_Test_Click    '【テスト】
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
           Case vbKeyEscape
                If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
                If PB_Can2.Visible Then Call PB_Can2_Click  '【テスト取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      '【入力】
                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
                                  Dummy.Enabled = False
                               End If
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
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
   Call CisFrm.MousePT(11)
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      iH1_Seban.SetFocus
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    追　加  キ ー(F1)        +
'+----------------------------+
Private Sub PB_Add_Click()
    If SyoriKB <> "ADD" Then
       K_Sykbnm = "追加"
       K_Sykbnm.ForeColor = PB_Add.BtnForeColor
       SyoriKB = "ADD"
       ProcHB = "H1"
       vsGrid1.HighLight = flexHighlightNever
       Call DispChange(ProcHB)
    End If
End Sub
'+----------------------------+
'+    削　除  キ ー(F3)        +
'+----------------------------+
Private Sub PB_Del_Click()
    If SyoriKB <> "DEL" Then
       K_Sykbnm = "削除"
       K_Sykbnm.ForeColor = PB_Del.BtnForeColor
       SyoriKB = "DEL"
       ProcHB = "B2"
       Call DispChange(ProcHB)
       vsGrid1.HighLight = flexHighlightAlways
       vsGrid1.SetFocus
    End If
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = "Select HT.* From 品番取引先マスタ HT"
    gSL_Select = gSL_Select & " Where HT.背番号 + HT.取引先 + HT.受入 + HT.品番 < '"
    gSL_Select = gSL_Select & RTrim(iH1_Seban) & RTrim(iH1_Torcd) & RTrim(iH1_Ukeir) & RTrim(iH1_Hinbn)
    gSL_Select = gSL_Select & "'"
    gSL_Select = gSL_Select & " And 品目 = 1 and SK区分 = 0 And 手配区分 = '2'"
    gSL_Select = gSL_Select & " ORDER BY HT.背番号 + HT.取引先 + HT.受入 + HT.品番 Desc "
    Call DBInput("Back")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Seban.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = "Select HT.* From 品番取引先マスタ HT"
    gSL_Select = gSL_Select & " Where HT.背番号 + HT.取引先 + HT.受入 + HT.品番 > '"
    gSL_Select = gSL_Select & RTrim(iH1_Seban) & RTrim(iH1_Torcd) & RTrim(iH1_Ukeir) & RTrim(iH1_Hinbn)
    gSL_Select = gSL_Select & "'"
    gSL_Select = gSL_Select & " And 品目 = 1 and SK区分 = 0 And 手配区分 = '2'"
    gSL_Select = gSL_Select & " ORDER BY HT.背番号 + HT.取引先 + HT.受入 + HT.品番 "
    Call DBInput("Next")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Seban.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    印　刷  キ ー(F8)       +
'+----------------------------+
Private Sub PB_Prt_Click()
'    kn_Call = "C"
'    kn_PrtKB = 1
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "   　かんばん発行を行いますか？       "
        .MB_Title = ""
        .MB_Button = Yes_No
        If Not .MBOX Then Exit Sub
    End With
    
    CKN3230_TestPrt = False
    
    If Not CKN3230LPrint Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　かんばん発行データがありません       "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Exit Sub
    End If
    Call GridSet
    
End Sub
'+----------------------------+
'+    テスト  キ ー(F9)       +
'+----------------------------+
Private Sub PB_Test_Click()
    CKN3230_TestPrt = True
    Call DispChange("B3")
    Call iB3_HSu.SetFocus
End Sub
Private Sub PB_ENT2_Click()
    CKN3230_TestPrt = True
    
   If Not Body3Chk Then Exit Sub
   If Not DBPut Then Exit Sub
   Call CKN3230LPrint
   
    CKN3230_TestPrt = False
    
    Call GridSet
   
   Call DispChange("B1")
End Sub
Private Sub PB_Can2_Click()
   Call DispChange("B1")
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Seban" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_UTorcd: Exit Sub
End Sub
'+----------------------------+
'+       仕入先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 6
    RV_Call = ""
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'+-----------------------------+
'+       取引先受入検索         +
'+-----------------------------+
Private Sub Look_UTorcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 6
    RV_Torcd = iH1_Torcd
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RV_Ukeir
    End If
End Sub
'+-----------------------------+
'+       品番検索              +
'+-----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
        
    RVI_Hinmoku = 1 ' 発注
    RVI_SK = 0      ' かんばん
    RV_Call = ""
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RV_Ukeir
        iH1_Hinbn = RV_Hinbn
        iH1_Seban = RV_Seban
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
    ' 背番号より他情報表示
    If Trim(iH1_Seban) <> "" Then
        gSL_Select = "Select 取引先,受入,品番 From 品番取引先マスタ"
        gSL_Select = gSL_Select & " Where 背番号 = '" & iH1_Seban & "'"
        gSL_Select = gSL_Select & "   And 品目 = 1"
        gSL_Select = gSL_Select & "   And SK区分 = 0"
        gSL_Select = gSL_Select & "   And 手配区分 = '2'"
        If HTMRead(gSL_Select, 0, 1) Then
           If CisDB.RecordCount = 1 Then
              iH1_Torcd = HTM.取引先
              iH1_Ukeir = HTM.受入
              iH1_Hinbn = HTM.品番
              Call iH1_Torcd_LostFocus
           End If
        End If
        Call HTMClose(1)
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
    If TorNmGet(iH1_Torcd, 6, 1) Then
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
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
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
    Dim B2Mode      As Boolean
    Dim B2Color     As Long
    Dim B3Mode      As Boolean
    Dim B3Color     As Long
    H1Mode = False
    H1Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
    B2Mode = False
    B2Color = gPLostSel
    B3Mode = False
    B3Color = gPLostSel
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
       Case "B2"
          B2Mode = True
          B2Color = gPGotSel
       Case "B3"
          B3Mode = True
          B3Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B1_Area2.Enabled = B1Mode
    B2_Area1.Enabled = Not B3Mode
    B2_Area1.BackColor = B2Color
    B3_Area1.Visible = B3Mode
    B3_Area1.Enabled = B3Mode
    B3_Area1.BackColor = B3Color
    
    PB_Look.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = B1Mode
    PB_Del.Visible = H1Mode
    PB_Add.Visible = B2Mode
    PB_Test.Visible = B1Mode
    
    PNL_PF.Visible = Not B3Mode
    
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
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    Call ImageDisp("", Image_Point, NoImage_Point)
    Call ImageDisp("", Image_Hinban, NoImage)
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B2" Then
        If Not DBDelete Then
           vsGrid1.SetFocus
           GoTo ReturnPress_Ed
        End If
        vsGrid1.SetFocus
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
    If Trim(iH1_Seban) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    背番号を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Seban.SetFocus
       Exit Function
    End If
    
    gSL_Select = "select * From 品番取引先マスタ"
    gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & " And   背番号 = '" & RTrim(iH1_Seban) & "'"
    If Not HTMRead(gSL_Select, 1, 1) Then
        gSL_Select = "Select HT.背番号,HT.取引先,HT.受入,HT.品番,HN.表示品番,HN.社内背番,"
        gSL_Select = gSL_Select & " HT.最終発行連番,Rtrim(TR.略称) 取引先名,HT.手配区分"
        gSL_Select = gSL_Select & " From 品番取引先マスタ HT"
        gSL_Select = gSL_Select & " Left Outer Join 品番マスタ HN"
        gSL_Select = gSL_Select & "   On HN.品番 = HT.品番"
        gSL_Select = gSL_Select & " Left Outer Join 取引先マスタ TR"
        gSL_Select = gSL_Select & "   On TR.取引先CD = HT.取引先"
        gSL_Select = gSL_Select & "   And TR.取引先区分 = 1"
        gSL_Select = gSL_Select & " Where HT.背番号 = '" & iH1_Seban & "'"
        gSL_Select = gSL_Select & "   And HT.品目 = 1"
        gSL_Select = gSL_Select & "   And HT.SK区分 = 0"
        gSL_Select = gSL_Select & "   And HT.手配区分 = '2'"
        gSL_Select = gSL_Select & " Order By HT.品番,HT.取引先,HT.受入"
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
           End If
        End If
        Call HTMClose(1)
    End If
    
    If Trim(iH1_Torcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    社内手配先を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    End If
    If Not TorNmGet(iH1_Torcd, 6) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    社内手配先　取引先マスタ未登録            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    End If
    If Trim(iH1_Ukeir) <> "" Then
        gSL_Select = "select * From 取引先受入マスタ "
        gSL_Select = gSL_Select & " WHERE 取引先CD = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & "   And 取引先区分 = 1"
        If Not TUMRead(gSL_Select, 1) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    受入未登録            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_Ukeir.SetFocus
           Exit Function
        End If
    End If
    If Trim(iH1_Hinbn) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    品番を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Hinbn.SetFocus
       Exit Function
    End If

' マスタ入力 SELECT文 ｾｯﾄ
    gSL_Select = " Select HT.* From 品番取引先マスタ HT"
    gSL_Select = gSL_Select & " Where HT.品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   HT.取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   HT.受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & " And   HT.背番号 = '" & RTrim(iH1_Seban) & "'"
    If Not DBInput("INV") Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    品番取引先マスタ未登録            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_Seban.SetFocus
        Exit Function
    Else
        If HTM.品目 <> 1 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     発注品番を指定して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Seban.SetFocus
            Exit Function
        End If
        If HTM.SK区分 <> 0 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     かんばんを指定して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Seban.SetFocus
            Exit Function
        End If
        If HTM.手配区分 <> "2" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     手配先が'社内'を指定して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Seban.SetFocus
            Exit Function
        End If
    End If
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Dim NoStr       As String
    Body1Chk = False
    
'    ' 追加枚数のみ（１）
    If Not Back_HSu(1).Visible And Not Back_SSu(0).Visible And Not Back_SSu(1).Visible Then
       If CisFun.Val2(iB1_HSu(0)) = 0 Then
          With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     追加枚数を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
          End With
          Exit Function
       End If
    
    End If
    
    If CisFun.Val2(iB1_HSu(0)) + CisFun.Val2(B1lb_EndSu) >= 1000 Then
       With CisFun
         .MB_Lines = 4
         .MB_MSG(2) = "     ９９９枚以上は発行出来ません            "
         .MB_Title = ""
         .MB_Button = Error
         .MBOX
       End With
       Exit Function
    End If
    If Back_SSu(1).Visible = True Then
        If CisFun.Val2(iB1_HSu(1)) + CisFun.Val2(B1lb_EndSu) >= 1000 Then
           With CisFun
             .MB_Lines = 4
             .MB_MSG(2) = "     ９９９枚以上は発行出来ません            "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
           End With
           Exit Function
        End If
    End If
    
    If Not Back_HSu(1).Visible And Back_SSu(0).Visible And Not Back_SSu(1).Visible Then
       If CisFun.Val2(iB1_HSu(0)) = 0 And _
          CisFun.Val2(iB1_SSSu(0)) = 0 And _
          CisFun.Val2(iB1_ESSu(0)) = 0 Then
          With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     追加枚数又は再発行№を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
          End With
          Exit Function
       End If
    End If
    If Trim(B1lb_SHYmd) <> "" Then
       If CisFun.Val2(iB1_HSu(0)) = 0 And _
          CisFun.Val2(iB1_SSSu(0)) = 0 And _
          CisFun.Val2(iB1_ESSu(0)) = 0 And _
          CisFun.Val2(iB1_HSu(1)) = 0 And _
          CisFun.Val2(iB1_SSSu(1)) = 0 And _
          CisFun.Val2(iB1_ESSu(1)) = 0 Then
          With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     追加枚数又は再発行№を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
          End With
          Exit Function
       End If
     End If
    For gInt = 0 To 1
        If gInt = 0 Then
           NoStr = "①"
        Else
           NoStr = "②"
        End If
        If Trim(B1lb_SHYmd) = "" Then
           NoStr = ""
        End If
        If Back_SSu(gInt).Visible And _
           CisFun.Val2(iB1_HSu(gInt)) = 0 And _
           CisFun.Val2(iB1_SSSu(gInt)) = 0 And _
           CisFun.Val2(iB1_ESSu(gInt)) <> 0 Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "     再発行№" & NoStr & "(開始)を入力して下さい            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB1_SSSu(gInt).SetFocus
           Exit Function
        End If
        If Back_SSu(gInt).Visible And _
           CisFun.Val2(iB1_HSu(gInt)) = 0 And _
           CisFun.Val2(iB1_ESSu(gInt)) = 0 And _
           CisFun.Val2(iB1_SSSu(gInt)) <> 0 Then
           With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "     再発行№" & NoStr & "(終了)を入力して下さい            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
            End With
            iB1_ESSu(gInt).SetFocus
            Exit Function
         End If
         If CisFun.Val2(iB1_SSSu(gInt)) > CisFun.Val2(iB1_ESSu(gInt)) Then
            With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "     再発行№" & NoStr & "を正しく入力して下さい（開始＞終了）            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
            End With
            iB1_ESSu(gInt).SetFocus
            Exit Function
         End If
         If CisFun.Val2(iB1_ESSu(gInt)) > CisFun.Val2(B1lb_EndSu) Then
            With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "     再発行№" & NoStr & "は最終№以下を指定して下さい            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
            End With
            iB1_ESSu(gInt).SetFocus
            Exit Function
         End If
    Next gInt
    
    Body1Chk = True
End Function
'****************************
'*      Body3ﾁｪｯｸ処理
'****************************
Private Function Body3Chk() As Boolean
    Dim NoStr       As String
    Body3Chk = False
    
    If CisFun.Val2(iB3_HSu) = 0 Then
       With CisFun
         .MB_Lines = 4
         .MB_MSG(2) = "     追加枚数を入力して下さい            "
         .MB_Title = ""
         .MB_Button = Error
         .MBOX
       End With
       Call iB3_HSu.SetFocus
       Exit Function
    End If
    
    Body3Chk = True
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
        iH1_Torcd = RTrim(HTM.取引先)
        iH1_Ukeir = RTrim(HTM.受入)
        iH1_Hinbn = RTrim(HTM.品番)
        Call BodySet("")
        DBInput = True
    End If
DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function

'**************************************
'*          BODY 内容ｾｯﾄ処理           *
'**************************************
Private Sub BodySet(SetType As String)
    If HTM.品目 <> 1 Or HTM.SK区分 <> 0 Then Exit Sub

    With HTM
        B1lb_SLot(0) = Format(.ロット1, "#,###")
        B1lb_SLot(1) = Format(.ロット2, "#,###")
        B1lb_Okiba = .置場
        
        B1lb_Syuki(0) = .収容器1
        B1lb_Syuyo(0) = Format(.収容数1, "#,###")
        B1lb_Syuki(1) = .収容器2
        B1lb_Syuyo(1) = Format(.収容数2, "#,###")
        
        If Trim(.変更日) <> "" Then
           B1lb_SHYmd = Mid(.変更日, 1, 4) & "/" & Mid(.変更日, 5, 2) & "/" & _
                        Mid(.変更日, 7, 2)
           B1lb_SHBin = .変更便
           Back_HSu(1).Visible = True
           Back_SSu(1).Visible = True
        Else
           Back_HSu(1).Visible = False
           Back_SSu(1).Visible = False
        End If
        If Not CKN3230_TestPrt Then
           If .最終発行連番 <> 0 Then
              B1lb_EndSu = Format(.最終発行連番, "000")
              Back_SSu(0).Visible = True
           Else
              Back_SSu(0).Visible = False
              Back_SSu(1).Visible = False
           End If
        End If
        
        ' 返却場所
        If MeisyoGet("返却場所", HTM.返却場所) Then
           B1lb_Henkyaku = SYM_Meisyo
        End If
        ' 組合せ場所
        gSL_Select = ""
        gSL_Select = gSL_Select & " SELECT * FROM 名称マスタ"
        gSL_Select = gSL_Select & "  WHERE 区分名称 = '工程かんばん発行パターン'"
        'gSL_Select = gSL_Select & "    AND 値 = '" & .後工程 & "'"
        gSL_Select = gSL_Select & "    AND 値 = '" & Format(.後工程, "000") & "'"
        gSL_Select = gSL_Select & "    AND 区分タイプ = ''"
        If SYMRead(gSL_Select, 1) Then
           If TorNmGet(SYM.英数字2, 6) Then
              B1lb_AKotei(0) = TRM_RName
           End If
           If MeisyoGet("返却場所", SYM.英数字3) Then
              B1lb_AKotei(1) = SYM_Meisyo
           End If
        End If
        
        ' ポイント
        B1lb_Point = .ポイント
        gSL_Select = ""
        gSL_Select = gSL_Select & " SELECT * FROM ポイントマスタ"
        gSL_Select = gSL_Select & "  WHERE ポイント = '" & HTM.ポイント & "'"
        If POMRead(gSL_Select, 1) Then
           Call ImageDisp(POM.イメージ, Image_Point, NoImage_Point)
        End If
        
    End With
    ' 品番情報
    gSL_Select = "select * From 品番マスタ "
    gSL_Select = gSL_Select & " Where 品番 = '" & HTM.品番 & "'"
    If HNMRead(gSL_Select, 1) Then
       With HNM
            B1lb_Syakei = .車型
            B1lb_Hinnm = .品名
            B1lb_SeiKu = HNM.ABC
        
            '-------------------- 品番
            Call ImageDisp(.IMAGE, Image_Hinban, NoImage)
            '得意先
            If TorNmGet(HNM.得意先, 3) Then
               B1lb_Tokui = TRM_RName
            End If
            Select Case .部品区分
                   Case 0: B1lb_BuKB = "完成品"
                   Case 1: B1lb_BuKB = "購入品"
                   Case 2: B1lb_BuKB = "仕掛品"
            End Select
       End With
    End If
    '
    Call iH1_Torcd_LostFocus
    
End Sub
'+--------------------------------------------+
'+      イメージ表示
'+--------------------------------------------+
Private Sub ImageDisp(IMAGE As String, ImageObj As IMAGE, NoImageObj As Object)
    Dim ImageCheckFlg   As Boolean
    
    Set ImageObj.Picture = LoadPicture("")
    NoImageObj.Left = ImageObj.Left + 30
    NoImageObj.Top = ImageObj.Top + 30
    NoImageObj.Visible = True
    
    If RTrim(IMAGE) = "" Then Exit Sub
    
    On Error GoTo ImageDisp_Ex
    ImageCheckFlg = False
    If Dir(RTrim(CKN3230_ImagePath) & IMAGE) <> "" Then
        Set ImageObj.Picture = LoadPicture(RTrim(CKN3230_ImagePath) & IMAGE)
        ImageCheckFlg = True
        NoImageObj.Visible = False
        DoEvents
    End If
ImageDisp_Ex:
End Sub
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+    削除 処理          +
'+----------------------+
Private Function DBDelete() As Boolean
    DBDelete = False
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     選択行を削除します      "
        .MB_Title = "削除処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
'
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

    With CisDB
        .SQL = "Delete From かんばん発行テーブル"
        .SQL = .SQL & " Where かんばん種類 = 1"
        .SQL = .SQL & " And   背番号 = '" & RTrim(vsGrid1.TextMatrix(vsGrid1.Row, 5)) & "'"
        .SQL = .SQL & " And   品番 = '" & RTrim(vsGrid1.TextMatrix(vsGrid1.Row, 9)) & "'"
        .SQL = .SQL & " And   取引先 = '" & RTrim(vsGrid1.TextMatrix(vsGrid1.Row, 1)) & "'"
        .SQL = .SQL & " And   受入 = '" & RTrim(vsGrid1.TextMatrix(vsGrid1.Row, 3)) & "'"
        .SQL = .SQL & " And   収容数 = " & CisFun.Val2(vsGrid1.TextMatrix(vsGrid1.Row, 7))
        If Not .DBExec Then GoTo DBDelete_Err
    End With
'
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

'
    Call GridSet
'
    DBDelete = True
    Exit Function
DBDelete_Err:
    CisDB.DBTran (TransRollback) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加 処理          +
'+----------------------+
Private Function DBPut() As Boolean
    DBPut = False
    Dim AddMai  As Long
    Dim PrtStart    As Long
    Dim PrtEnd      As Long
    Dim wTorNm      As String
    Dim wwTorcd     As String
    Dim wNTorNm     As String
    Dim PrtKb       As Integer
    Dim wPutFlg     As Boolean
    
    If Not CKN3230_TestPrt Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     発行データ作成処理を行います      "
            .MB_Title = "処理確認"
            .MB_MSG(3) = "     よろしいですか？      "
            .MB_Button = OK_CAN
            If Not .MBOX Then Exit Function
        End With
    End If
    On Error GoTo DBPut_Err
'   品番取引先情報
    gSL_Select = "Select *"
    gSL_Select = gSL_Select & " From 品番取引先マスタ "
    gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & " And   背番号 = '" & RTrim(iH1_Seban) & "'"
    gSL_Select = gSL_Select & " And   品目 = 1"
    gSL_Select = gSL_Select & " And   SK区分 = 0"
    If Not HTMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     品番取引先マスタ　異常      "
            .MB_Title = ""
            .MB_Button = Error
            Call .MBOX
        End With
        Exit Function
    End If
    ' 品番情報
    gSL_Select = "Select * From 品番マスタ "
    gSL_Select = gSL_Select & " Where 品番 = '" & HTM.品番 & "'"
    If Not HNMRead(gSL_Select, 1) Then
       Call ItemsClearHNM
    End If
    ' 取引先情報
    gSL_Select = "Select * From 取引先マスタ "
    gSL_Select = gSL_Select & " Where 取引先CD = '" & HTM.取引先 & "'"
    gSL_Select = gSL_Select & "   And 取引先区分 = 1"
    If Not TRMRead(gSL_Select, 1) Then
       Call ItemsClearTRM
    End If
    wTorNm = TRM.略称
    wwTorcd = TRM.取引先
    If Trim(TRM.工場CD) <> "" Then
       wwTorcd = wwTorcd & "-" & TRM.工場CD
    End If
    ' 納入先情報
    wNTorNm = ""
    If Trim(HTM.納入先) <> "" Then
        gSL_Select = "Select * From 取引先マスタ "
        gSL_Select = gSL_Select & " Where 取引先CD = '" & HTM.納入先 & "'"
        gSL_Select = gSL_Select & "   And 取引先区分 = 1"
        If Not TRMRead(gSL_Select, 1) Then
           Call ItemsClearTRM
        End If
        wNTorNm = TRM.略称
    End If
    ' ポイント情報
    gSL_Select = "Select * From ポイントマスタ "
    gSL_Select = gSL_Select & " Where ポイント= '" & HTM.ポイント & "'"
    If Not POMRead(gSL_Select, 1) Then
       Call ItemsClearPOM
    End If
    
    ' かんばん発行情報出力
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' 発行済データ削除
    With CisDB
        .SQL = "Delete かんばん発行テーブル "
        .SQL = .SQL & "    Where かんばん種類 = 1"
        .SQL = .SQL & "      And 発行済 = 1"
        .DBExec
    End With
' ( 追加発行 )
    For gInt = 0 To 1
        PrtStart = 0: PrtEnd = 0
        If Not CKN3230_TestPrt Then
           AddMai = CisFun.Val2(iB1_HSu(gInt))
           If AddMai <> 0 Then
              PrtStart = CisFun.Val2(B1lb_EndSu) + 1
              PrtEnd = PrtStart + AddMai - 1
              PrtKb = 0        ' 追加発行
           End If
           If CisFun.Val2(iB1_SSSu(gInt)) <> 0 Then
              PrtStart = CisFun.Val2(iB1_SSSu(gInt))
              PrtKb = 1        ' 再発行
           End If
           If CisFun.Val2(iB1_ESSu(gInt)) <> 0 Then
              PrtEnd = CisFun.Val2(iB1_ESSu(gInt))
              PrtKb = 1        ' 再発行
           End If
        Else
           If gInt = 0 Then
              AddMai = CisFun.Val2(iB3_HSu)
              PrtStart = CisFun.Val2(B1lb_EndSu) + 1
              PrtEnd = PrtStart + AddMai - 1
              PrtKb = 0        ' 追加発行
           End If
        End If
        If PrtEnd = 0 Then
           PrtEnd = PrtStart
        End If
        
        If PrtStart <> 0 Then
            For gLong = PrtStart To PrtEnd
                wPutFlg = False
                gSL_Select = "Select * From かんばん発行テーブル"
                gSL_Select = gSL_Select & " Where かんばん種類 = 1"
                gSL_Select = gSL_Select & " And   背番号 = '" & RTrim(HTM.背番号) & "'"
                gSL_Select = gSL_Select & " And   品番 = '" & RTrim(HTM.品番) & "'"
                gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(HTM.取引先) & "'"
                gSL_Select = gSL_Select & " And   受入 = '" & RTrim(HTM.受入) & "'"
                gSL_Select = gSL_Select & " And   発行NO = " & gLong
                If Not KHTRead(gSL_Select, 1) Then
                    Call ItemsClearKHT
                    With KHT
                        .かんばん種類 = 1
                        .品番 = HTM.品番
                        .取引先 = HTM.取引先
                        .受入 = HTM.受入
                        .発行NO = gLong
                    End With
                    wPutFlg = True
                End If
                With KHT
                    .背番号 = HTM.背番号
                    If Not CKN3230_TestPrt Then
                       .発行区分 = PrtKb
                    Else
                       .発行区分 = 3
                    End If
                    .発行種類 = 1
                    .表示品番 = HNM.表示品番
                    If Len(HNM.表示品番) < 15 Then
                       .品番14 = HNM.表示品番
                    Else
                       .品番30 = HNM.表示品番
                    End If
                    .品目 = 3 'HTM.品目
                    .品名 = HNM.品名
                    .受入名 = TUM.受入名称
                    .収容器 = B1lb_Syuki(gInt)
                    .収容数 = CisFun.Val2(B1lb_Syuyo(gInt))
                    .収容区分 = gInt + 1
                    .手配区分 = HTM.手配区分
                    .手配区分名 = "社内"
                    .アドレス = HTM.アドレス
                    .IMAGE = HNM.IMAGE
                    .取引先CD = wwTorcd
                    .取引先名 = TRM.略称
                    .取引先名1 = CisFun.Mid2(TRM.略称, 1, 10)
                    .取引先名2 = CisFun.Mid2(TRM.略称, 11, 10)
                    .取引先工場名 = TRM.工場名
                    .取引先工場名1 = CisFun.Mid2(TRM.工場名, 1, 10)
                    .取引先工場名2 = CisFun.Mid2(TRM.工場名, 11, 10)
                    If Trim(HTM.サイクル) <> "" Then
                        .サイクル = CisFun.RSetFld(Mid(HTM.サイクル, 1, 1), 1, "#") & "-" & _
                                    CisFun.RSetFld(Mid(HTM.サイクル, 2, 2), 2, "#") & "-" & _
                                    CisFun.RSetFld(Mid(HTM.サイクル, 4, 2), 2, "#")
                    End If
                    .置場 = HTM.置場
                    .車型 = HNM.車型
'                    .自社CD = Kn_Jisya
'                    .自社名1 = CisFun.Mid2(Kn_JisyaNM, 1, 20)
'                    .自社名2 = CisFun.Mid2(Kn_JisyaKJNM, 1, 14)
                    .実数区分 = HNM.実数区分
                    .備考 = HTM.備考
                    .ポイント = HTM.ポイント
                    .ポイントIMAGE = CKN3230_ImagePath & POM.イメージ
                    .返却場所 = HTM.返却場所
                    .返却場所名 = B1lb_Henkyaku
                    .後工程 = Format(HTM.後工程, "000")
                    .後工程名1 = B1lb_AKotei(0)
                    .後工程名2 = B1lb_AKotei(1)
                    .基準工数 = HTM.基準工数
'                    .基準生産日 = B1lb_SKijyun
                    .生産ロット = CisFun.Val2(B1lb_SLot(gInt))
                    .生産区分 = HNM.ABC
                    .収容区分 = gInt + 1
                    .IMAGE = CKN3230_ImagePath & HNM.IMAGE
                    .得意先 = HNM.得意先
                    .得意先名 = B1lb_Tokui
                    .作成者 = gTanto
                    .部署名 = H1lb_Tornm
                    .部品区分 = HTM.部品区分
                    
                    Select Case HTM.部品区分
                           Case 0: .かんばん枠色 = mWakuColor0
                           Case 1: .かんばん枠色 = mWakuColor1
                           Case 2: .かんばん枠色 = mWakuColor2
                    End Select
                    
                End With
            
                If wPutFlg Then
                    Call KHTInsert
                Else
                    Call KHTUpdate
                End If
                If Not CKN3230_TestPrt Then
                   B1lb_EndSu = Format(gLong, "000")
                End If
            Next gLong
        End If
        
        If AddMai <> 0 And Not CKN3230_TestPrt Then
           With CisDB
                .SQL = "Update 品番取引先マスタ Set  "
                .SQL = .SQL & "最終発行連番 = " & PrtEnd & ","
                .SQL = .SQL & "更新日 = GETDATE(),"
                .SQL = .SQL & "更新者 = '" & gTanto & "',"
                .SQL = .SQL & "更新端末 = SubString(Host_Name(),1,20)"
                .SQL = .SQL & " Where 品番 = '" & RTrim(HTM.品番) & "'"
                .SQL = .SQL & "   And 取引先 = '" & RTrim(HTM.取引先) & "'"
                .SQL = .SQL & "   And 受入 = '" & RTrim(HTM.受入) & "'"
                If Not .DBExec Then GoTo DBPut_Err
           End With
           Back_SSu(gInt).Enabled = True
        End If
    Next gInt

    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    Call GridSet
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisDB.DBTran (TransRollback) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    CisFun.ErrorBox
    End
End Function
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 0, 0, "№")
        Call .FixedSet(中中, 左中, あり, 7, 1, "手配先")
        Call .FixedSet(中中, 左中, あり, 22, 2, "手配先名(社内)")
        Call .FixedSet(中中, 中中, あり, 2, 3, "受")
        Call .FixedSet(中中, 左中, あり, 25, 4, "品　番")
        Call .FixedSet(中中, 左中, あり, 6, 5, "背番号")
        Call .FixedSet(中中, 左中, あり, 10, 6, "収容器")
        Call .FixedSet(中中, 右中, あり, 7, 7, "収容数")
        Call .FixedSet(中中, 右中, あり, 0, 8, "発行枚数")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 9, "編集前品番")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8")
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
    
    CisVsGrid.Clear

' 抽出処理 初期ｾｯﾄ
    gSL_Select = "select KT.取引先,KT.受入,KT.表示品番,KT.背番号,KT.収容器,KT.品番,"
    gSL_Select = gSL_Select & "KT.収容数,Count(KT.発行NO) 件数,"
    gSL_Select = gSL_Select & "    TR.略称 取引先名1"
    gSL_Select = gSL_Select & "       from かんばん発行テーブル KT"
    gSL_Select = gSL_Select & "    Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & "    On TR.取引先CD = KT.取引先"
    gSL_Select = gSL_Select & "    And TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "    Where KT.かんばん種類 = 1"
    gSL_Select = gSL_Select & "      And KT.発行済 = 0"
    gSL_Select = gSL_Select & "    Group by KT.取引先,KT.受入,KT.表示品番,KT.背番号,KT.収容器,KT.品番,"
    gSL_Select = gSL_Select & "             KT.収容数,TR.略称"
    gSL_Select = gSL_Select & "    Order by MAX(KT.作成日) Desc"
    If Not KHTRead(gSL_Select) Then GoTo GridSet_Ed
    
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid1
       .Redraw = False
        If CisDB.RecordCount + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = CisDB.RecordCount + 1
        End If
    End With
    ID = 0: Ttl_Kensu = 0
    Do Until Not KHT_RDSTS
        With vsGrid1
            ID = ID + 1
            '№
            .TextMatrix(ID, 0) = Format(ID, "####")
            .TextMatrix(ID, 1) = KHT.取引先
            .TextMatrix(ID, 2) = KHT.取引先名1
            .TextMatrix(ID, 3) = KHT.受入
            .TextMatrix(ID, 4) = KHT.表示品番
            .TextMatrix(ID, 5) = KHT.背番号
            .TextMatrix(ID, 6) = KHT.収容器
            .TextMatrix(ID, 7) = Format(KHT.収容数, "#,###")
            .TextMatrix(ID, 8) = Format(KHT.件数, "###")
            .TextMatrix(ID, 9) = KHT.品番
            Ttl_Kensu = Ttl_Kensu + KHT.件数
       End With
       Call KHTReadNext
    Loop
    With vsGrid1
         iT1_Addsu = Format(Ttl_Kensu, "#,###")
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
GridSet_Ed:
    Call KHTClose
End Sub

Private Sub vsGrid1_DblClick()
    If ProcHB = "B2" Then
       If Trim(vsGrid1.TextMatrix(vsGrid1.Row, 1)) = "" Then Exit Sub
       Call ReturnPress
    End If
End Sub

Private Sub ColorGet()
' かんばん色
    With SYM
        Call ItemsClearSYMCreate
        With SYMCreate
            .区分名称 = "かんばん色設定"
            .区分桁数 = 1
            .値 = "0"
            .数字1 = KanbanColor0.BackColor
            .値名称 = "完成品"
            .作成者 = gTanto
            If Not SYMInitCreate(True) Then
                .値 = "1"
                .値名称 = "購入品"
                .数字1 = KanbanColor1.BackColor
                Call SYMInitCreate
                .値 = "2"
                .値名称 = "仕掛品"
                .数字1 = KanbanColor2.BackColor
                Call SYMInitCreate
            End If
        End With
        
        If MeisyoGet("かんばん色設定", "0") Then
           mWakuColor0 = .数字1
        End If
        If MeisyoGet("かんばん色設定", "1") Then
           mWakuColor1 = .数字1
        End If
        If MeisyoGet("かんばん色設定", "2") Then
           mWakuColor2 = .数字1
        End If
    End With
End Sub


