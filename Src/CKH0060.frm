VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1BA4D453-0150-11CE-89B0-0000C037528B}#3.2#0"; "PDQCom32.OCX"
Begin VB.Form CKH0060 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "発注かんばん入力 [一括] 手差し"
   ClientHeight    =   10530
   ClientLeft      =   210
   ClientTop       =   570
   ClientWidth     =   15150
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
   ScaleWidth      =   15150
   WindowState     =   2  '最大化
   Begin PdqcommLib.PDQComm PDQComm1 
      Height          =   480
      Left            =   45
      TabIndex        =   3
      Top             =   1425
      Width           =   480
      _Version        =   196610
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.26
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FastScroll      =   0   'False
      Object.Height          =   -20
      Object.Width           =   0
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   7080
      Left            =   420
      ScaleHeight     =   7020
      ScaleWidth      =   14505
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3450
      Width           =   14565
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   6900
         Left            =   60
         Top             =   60
         Width           =   14340
         _ExtentX        =   25294
         _ExtentY        =   12171
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
            Height          =   6810
            Index           =   1
            Left            =   45
            Top             =   45
            Width           =   14235
            _ExtentX        =   25109
            _ExtentY        =   12012
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
            Begin vsFlexLib.vsFlexArray vsGrid1 
               Height          =   6690
               Left            =   60
               TabIndex        =   2
               TabStop         =   0   'False
               Top             =   60
               Width           =   14115
               _Version        =   196608
               _ExtentX        =   24897
               _ExtentY        =   11800
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
               FormatString    =   "背番号|枚数|背番号|枚数|背番号|枚数|背番号|枚数|背番号|枚数|背番号|枚数|背番号|枚数|背番号|枚数|背番号|枚数|背番号|枚数|"
               Rows            =   11
               Cols            =   21
               FixedCols       =   0
               BackColor       =   16777152
               BackColorBkg    =   12632256
               FocusRect       =   0
               HighLight       =   0
               ScrollBars      =   2
               RowHeightMin    =   600
               GridLines       =   2
               SelectionMode   =   1
               BackColorAlternate=   16777152
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
      Caption         =   "【 発注かんばん入力　[一括] 手差し 】"
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
      Left            =   420
      Top             =   2760
      Width           =   6015
      _ExtentX        =   10610
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
         Width           =   1980
         _ExtentX        =   3493
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
            Width           =   840
            _ExtentX        =   1482
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
         Left            =   2010
         Top             =   30
         Width           =   1980
         _ExtentX        =   3493
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
            Width           =   855
            _ExtentX        =   1508
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
         Left            =   3990
         Top             =   30
         Width           =   1980
         _ExtentX        =   3493
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
            Width           =   840
            _ExtentX        =   1482
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
      Height          =   780
      Left            =   375
      Top             =   450
      Visible         =   0   'False
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   1376
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
   Begin Cis3D_v60.CIS3D Read_Msg 
      Height          =   1305
      Left            =   465
      Top             =   1305
      Visible         =   0   'False
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   2302
      BackColor       =   16777088
      Caption         =   "読み取り中"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   -2147483643
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   8
      cAlingnment     =   7
   End
   Begin Cis3D_v60.CIS3D PNL_PF 
      Height          =   855
      Left            =   11220
      Top             =   870
      Width           =   3765
      _ExtentX        =   6641
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
         Left            =   1875
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
         Left            =   75
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
   End
   Begin Cis3D_v60.CIS3D CIS3D1 
      Height          =   1185
      Left            =   8865
      Top             =   2205
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   2090
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
      Begin Cis3D_v60.CIS3D B1_ErrName 
         Height          =   570
         Index           =   2
         Left            =   4080
         Top             =   30
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   1005
         BackColor       =   16761087
         ForeColor       =   0
         Caption         =   "抜き取り"
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
         Begin Cis3D_v60.CIS3D B1lb_ErrCnt 
            Height          =   360
            Index           =   2
            Left            =   1300
            Top             =   100
            Width           =   650
            _ExtentX        =   1138
            _ExtentY        =   635
            BackColor       =   14737632
            Caption         =   "###"
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
      Begin Cis3D_v60.CIS3D B1_ErrName 
         Height          =   570
         Index           =   0
         Left            =   30
         Top             =   30
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   1005
         BackColor       =   16761087
         ForeColor       =   0
         Caption         =   "重複読取"
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
         Begin Cis3D_v60.CIS3D B1lb_ErrCnt 
            Height          =   360
            Index           =   0
            Left            =   1300
            Top             =   100
            Width           =   650
            _ExtentX        =   1138
            _ExtentY        =   635
            BackColor       =   14737632
            Caption         =   "###"
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
      Begin Cis3D_v60.CIS3D B1_ErrName 
         Height          =   570
         Index           =   1
         Left            =   2055
         Top             =   30
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   1005
         BackColor       =   16761087
         ForeColor       =   0
         Caption         =   "ﾁｪｯｸｴﾗｰ"
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
         Begin Cis3D_v60.CIS3D B1lb_ErrCnt 
            Height          =   360
            Index           =   1
            Left            =   1300
            Top             =   100
            Width           =   650
            _ExtentX        =   1138
            _ExtentY        =   635
            BackColor       =   14737632
            Caption         =   "###"
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
      Begin Cis3D_v60.CIS3D B1_ErrName 
         Height          =   570
         Index           =   3
         Left            =   30
         Top             =   600
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   1005
         BackColor       =   16761087
         ForeColor       =   0
         Caption         =   "打切(端数)"
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
         Begin Cis3D_v60.CIS3D B1lb_ErrCnt 
            Height          =   360
            Index           =   3
            Left            =   1300
            Top             =   100
            Width           =   650
            _ExtentX        =   1138
            _ExtentY        =   635
            BackColor       =   14737632
            Caption         =   "###"
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
      Begin Cis3D_v60.CIS3D B1_ErrName 
         Height          =   570
         Index           =   4
         Left            =   2055
         Top             =   600
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   1005
         BackColor       =   16761087
         ForeColor       =   0
         Caption         =   "打切(完了)"
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
         Begin Cis3D_v60.CIS3D B1lb_ErrCnt 
            Height          =   360
            Index           =   4
            Left            =   1300
            Top             =   100
            Width           =   650
            _ExtentX        =   1138
            _ExtentY        =   635
            BackColor       =   14737632
            Caption         =   "###"
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
      Begin Cis3D_v60.CIS3D B1_ErrName 
         Height          =   570
         Index           =   5
         Left            =   4080
         Top             =   600
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   1005
         BackColor       =   16761087
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   50
         Begin Cis3D_v60.CIS3D B1lb_ErrCnt 
            Height          =   360
            Index           =   5
            Left            =   1300
            Top             =   100
            Width           =   650
            _ExtentX        =   1138
            _ExtentY        =   635
            BackColor       =   14737632
            Caption         =   "###"
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
End
Attribute VB_Name = "CKH0060"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   発注かんばん入力[一括]手差し
'**       フォームID    :   CKH0060
'**       処理概要      :       ( TX-1609 )
'**
'**       作  成  日    :   2004/02/16  By CIS SIMPR-A
'**       変  更  日    :   2004/05/24  便算出時、カレンダ無しエラーが
'**                                     『EXECUTE後のトランザクション数は、COMMIT TRANまたは･･･』となるのを
'**                                         カレンダ未登録エラーに変更( ERROR No = 6 )
'**       変  更  日    :   2005/03/07  By CIS エラーメッセージ追加 .. 51:打切（端数）
'**       変  更  日    :   2006/02/03  By CIS かんばんエラー内容取得先変更
'**                                             (かんばんエラーポケット⇒かんばんエラー設定発注)
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim mSaveRow       As Integer       ' Row退避
    Dim mSaveCol       As Integer       ' Col退避
    
    Dim mKey            As String

    Dim mRId            As Long
    Dim mCId            As Long

    Dim mSColor         As Long
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
        
    mSColor = &HC0FFC0
    mMColor = &HFFFFC0
    mNColor = &HFFC0C0
    
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
    If Mid(Command, 1, 6) = "TX1301" Then
       TX1609_TX1301 = True
    Else
       TX1609_TX1301 = False
    End If
    Set TX1609_Comm = PDQComm1
    Call TX1609InfoGet
    
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
    Call GridClear
    
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
    Call TX1609Close
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
    PNL_PF.Visible = B1Mode Or RDMode
    PB_ENT.Visible = B1Mode
    
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
'+---------------------------------------+
'+      初期設定・初期表示
'+---------------------------------------+
Private Function Init_Set() As Boolean
    Init_Set = False
'<< かんばん読取エラーポケット表示 >>
    gSL_Select = " SELECT * FROM 名称マスタ"
'------------------------------------------------------------------------ 2006/02/03 Delete
'    gSL_Select = gSL_Select & " WHERE 区分名称 = 'かんばんエラーポケット'"
'------------------------------------------------------------------------ 2006/02/03 Delete End
'------------------------------------------------------------------------ 2006/02/03 Insert
    gSL_Select = gSL_Select & " WHERE 区分名称 = 'かんばんエラー設定発注'"
'------------------------------------------------------------------------ 2006/02/03 Insert End
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & " ORDER BY 値"
    If SYMRead(gSL_Select, 0) Then
        Do Until Not SYM_RDSTS
            Select Case SYM.値
                Case "02": B1_ErrName(0) = RTrim(SYM.値名称)
                Case "03": B1_ErrName(1) = RTrim(SYM.値名称)
                Case "04": B1_ErrName(2) = RTrim(SYM.値名称)
                Case "07": B1_ErrName(3) = RTrim(SYM.値名称)
                Case "06": B1_ErrName(4) = RTrim(SYM.値名称)
                Case "08": B1_ErrName(5) = RTrim(SYM.値名称)
            End Select
            Call SYMReadNext
        Loop
    End If
    Call SYMClose
    
    Init_Set = True
End Function
'//////////////////////////////////////////
'//
'//     かんばん読取
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
    If Not TX1609Open Then
        Call PB_END_Click
    End If
' デモ用設定
    kBas_DemoFno = 0
'///  かんばん読取処理 開始  １枚目
    Call ReadStart
    
    If kBas_ExecMode Then
       Do Until Not GetDemoData
          Call PDQComm1_OnComm
       Loop
    End If
    Exit Sub
'///
KanbanRead_Ed:
' 通信ﾎﾟｰﾄｸﾛｰｽﾞ
    Call TX1609Close
    
    Read_Msg.Visible = False
    ProcHB = "B1"
    Call DispChange(ProcHB)
    Call CisFrm.MousePT(1)
End Sub
'****************************************
'*  かんばん読取処理 開始  １枚目         *
'****************************************
Private Sub ReadStart()
    Read_Msg.Visible = True
' 初期処理
    Call TX1609Init
' 読取
    Call TX1609Read
    
    kBas_ReadFlg = True
    kBas_CommData = ""
    
'    Do Until Not kBas_ReadFlg           ' 2枚目移行
'       DoEvents
'    Loop
End Sub
'****************************************
'*  かんばん読取処理 開始  ２枚目以降     *
'****************************************
Private Sub PDQComm1_OnComm()
    Static ProcSTS  As Integer
    
    Call TX1609Event    ' 通信状態ﾁｪｯｸ

'   * エラー *
    If kBas_Event = "ERR" Then
         Call TX1609Close
         lb_ErrMsg = kBas_EventMsg
         lb_ErrMsg.Visible = True
         kBas_ReadFlg = True
         Call ReadEnd
         Exit Sub            ' 通信エラー処理
    End If
    
    If kBas_Event <> "RCV" Then Exit Sub   ' 受信ﾃﾞｰﾀ編集
    
    If Not TX1609SendProc Then Exit Sub    ' CR が存在しない時抜ける

    ' 電源ON
    If InStr(1, kBas_CommData, Chr(&H1B) & "P" & vbCr) > 0 Then GoTo ReadNext_ED

    If Mid(kBas_CommData, 1, 1) = Chr(&H1B) Then
       ' 読取終了
       If Mid(kBas_CommData, 2, 1) = "E" Then
          Call ReadEnd
          Exit Sub            ' 通信エラー処理
       End If
       ' キックバック(読取不能）
       If Mid(kBas_CommData, 2, 1) = "?" Then
          lb_ErrMsg = "読取不能"
          GoTo ReadNext_BK
       End If
       Read_Msg.BackColor = &HFFFF80
       ' 通信エラー(C)
       If Mid(kBas_CommData, 2, 1) = "C" Then
'          lb_ErrMsg = "通信エラー"
          GoTo ReadNext_ED
       End If
       ' 走行ｴﾗｰ(U)
       If Mid(kBas_CommData, 2, 1) = "U" Then
          lb_ErrMsg = "走行エラー"
          GoTo ReadNext_ED
       End If
       ' ﾀﾞﾌﾞﾙﾌｨｰﾄﾞｴﾗｰ(D)
       If Mid(kBas_CommData, 2, 1) = "D" Then
          lb_ErrMsg = "ダブルフィードエラー"
          GoTo ReadNext_ED
       End If
    End If
    
    lb_ErrMsg.Visible = False
    Read_Msg = "読み取り中"
    
'   かんばん内容チェック & 表示
    Call cKanbanValueSetT
    If QRArea.ステータス = 2 Then
       lb_ErrMsg = "対象外かんばん"
       GoTo ReadNext_BK
    End If
    Call KanbanProc
' 次かんばん読取
    Call TX1609Read
    DoEvents
    
    kBas_CommData = ""
    kBas_ReadFlg = False
    
    Exit Sub

ReadNext_ED:
    kBas_ReadFlg = False
    Call TX1609Init
    Call TX1609Read
    DoEvents
    Exit Sub
ReadNext_BK:
    Call TX1609KickBack
    Call TX1609Read
    Read_Msg = "読取エラー"
    lb_ErrMsg.Visible = True
    DoEvents
    kBas_ReadFlg = False
End Sub
'*******************************
'*  かんばん読取処理 終了        *
'*******************************
Private Sub ReadEnd()
' 読取中ﾒｯｾｰｼﾞ消去
    Read_Msg.Visible = False
    DoEvents
' 読取確認
    kBas_ReadMsg = "ReadEnd"

' 通信ポートクローズ
    Call TX1609Close

    ProcHB = "B1"
    Call DispChange(ProcHB)
    Read_Msg.Visible = False
    Call CisFrm.MousePT(1)

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
       kBas_ErrNo = QRArea.ステータス + 90
       kBas_ErrMsg = "その他エラー"
       GoTo KanbanProc_Set:
    End If
    
    mKey = Trim(QRArea.背番号)
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
    
KanbanProc_Set:
' ｸﾞﾘｯﾄﾞへ内容セット & 枚数カウント
    lb_InCnt = CisFun.Val2(lb_InCnt) + 1
    If kBas_ErrNo <> 0 And kBas_Poket > 0 And kBas_Poket < 9 Then
        lb_ErrCnt = Format(CisFun.Val2(lb_ErrCnt) + 1, "#,###")
        Select Case kBas_Poket
            Case 2: gInt = 0
            Case 3: gInt = 1
            Case 4: gInt = 2
            Case 7: gInt = 3
            Case 6: gInt = 4
            Case 8: gInt = 5
        End Select
        If kBas_Poket > 1 Then
           B1lb_ErrCnt(gInt) = Format(CisFun.Val2(B1lb_ErrCnt(gInt)) + 1, "###")
        End If
        Select Case kBas_ErrNo
            Case 1: gStr = "対象外かんばん(かんばん区分)"
            Case 2: gStr = "品番取引先未登録"
            Case 3: gStr = "発注サイクルマスタ未登録"
            Case 4: gStr = "かんばん流動管理ﾃｰﾌﾞﾙ未登録"
            Case 5: gStr = "かんばん重複読取"
            
'--------------------------------------------------< 2005.02.24 ADD START >
            Case 6: gStr = "カレンダマスタ未登録"
'--------------------------------------------------< 2005.02.24 ADD  END  >
            
            Case 10: gStr = "収容数相違"
            Case 11: gStr = "背番号相違"
            Case 15: gStr = "無効かんばん"
            Case 16: gStr = "未流動かんばん"
            Case 17: gStr = "滞留かんばん"
            Case 18: gStr = "打切り済"
            Case 19: gStr = "有効期限エラー"
            Case 20: gStr = "サイクルエラー"
            Case 21: gStr = "減かんばん(無効)"
            Case 99: gStr = kBas_ErrMsg
'▽▽:::::::::::::::::::::::::::::::::::::::::::::::::::::::<< 2005/03/07 >> Insert Start ::::::::::
            Case 51: gStr = "打切(端数)"
'△△:::::::::::::::::::::::::::::::::::::::::::::::::::::::<< 2005/03/07 >> Insert End   ::::::::::
        End Select
        lb_ErrMsg = gStr
        lb_ErrMsg.Visible = True
        Read_Msg = "読取エラー"
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "     " & gStr & "            "
'            .MB_Title = ""
'            .MB_Button = Error
'            .MBOX
'        End With
        Exit Function
    Else
        lb_OKCnt = Format(CisFun.Val2(lb_OKCnt) + 1, "#,###")
        Call GridSet(QRArea.背番号)
    End If

End Function

'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
Private Sub GridSet(SetItem As String)
    Dim wSetFlg     As Boolean
    
    With vsGrid1
        .Redraw = False
        .Cols = 20
        For mRId = 1 To .Rows - 1
            .Row = mRId
            For mCId = 1 To .Cols - 1 Step 2
                .Col = mCId
                .CellBackColor = mMColor
            Next mCId
        Next mRId
        
        wSetFlg = False
        For mRId = 1 To .Rows - 1
            .Row = mRId
            For mCId = 0 To .Cols - 1 Step 2
                If Trim(.TextMatrix(mRId, mCId)) = Trim(SetItem) Then
                   .TextMatrix(mRId, mCId + 1) = CisFun.Val2(.TextMatrix(mRId, mCId + 1)) + 1
                   .Col = mCId + 1: .CellBackColor = mNColor
                   wSetFlg = True
                   Exit For
                End If
                If Trim(.TextMatrix(mRId, mCId)) = "" And CisFun.Val2(.TextMatrix(mRId, mCId)) = 0 Then
                   .TextMatrix(mRId, mCId) = SetItem
                   .TextMatrix(mRId, mCId + 1) = 1
                   .Col = mCId + 1: .CellBackColor = mNColor
                   wSetFlg = True
                   Exit For
                End If
            Next mCId
            If wSetFlg Then Exit For
        Next mRId
        If Not wSetFlg Then
           .Rows = .Rows + 1
           .TextMatrix(.Rows - 1, 0) = SetItem
           .TextMatrix(.Rows - 1, 1) = 1
           .Row = .Rows - 1: .Col = 1
           .CellBackColor = mNColor
           For mCId = 0 To .Cols - 1
               .Col = mCId
                If mCId Mod 2 = 0 Or mCId = 0 Then
                   .CellBackColor = mSColor
                   .CellAlignment = flexAlignLeftCenter
                Else
                   .CellBackColor = mMColor
                   .CellAlignment = flexAlignRightCenter
                End If
           Next mCId
        End If
        .Redraw = True
    End With
End Sub
'**************************************
'*          グリッド　初期化           *
'**************************************
Private Sub GridClear()
    With vsGrid1
        .Redraw = False
        .Rows = 11: .Cols = 20
        For mRId = 1 To .Rows - 1
            .Row = mRId
            For mCId = 0 To .Cols - 1
                .Col = mCId
                If mCId Mod 2 = 0 Or mCId = 0 Then
                   .CellBackColor = mSColor
                   .CellAlignment = flexAlignLeftCenter
                Else
                   .CellBackColor = mMColor
                   .CellAlignment = flexAlignRightCenter
                End If
                .TextMatrix(mRId, mCId) = ""
            Next mCId
        Next mRId
        .Redraw = True
    End With
End Sub
