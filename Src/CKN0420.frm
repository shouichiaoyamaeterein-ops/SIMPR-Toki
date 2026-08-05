VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form CKN0420 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "かんばん流動管理問合せ【カレンダー】"
   ClientHeight    =   11325
   ClientLeft      =   3090
   ClientTop       =   3570
   ClientWidth     =   15330
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
   ScaleHeight     =   11325
   ScaleWidth      =   15330
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D PNL_Wait 
      Height          =   930
      Left            =   5070
      Top             =   4380
      Visible         =   0   'False
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   1640
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      Begin Cis3D_v60.CIS3D Message 
         Height          =   525
         Left            =   240
         Top             =   210
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   926
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "【抽出中】しばらくお待ち下さい"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         cBoderWidth     =   2
         cAlingnment     =   7
         cBoderStyle     =   1
      End
   End
   Begin Cis3D_v60.CIS3D Back_DispNo 
      Height          =   885
      Left            =   6120
      Top             =   7560
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   1561
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
      Begin Cis3D_v60.CIS3D Back_No 
         Height          =   465
         Left            =   1350
         Top             =   90
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   820
         Caption         =   "解除"
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
         cBoderWidth     =   4
         cAlingnment     =   7
         cButton         =   -1  'True
      End
      Begin Cis3D_v60.CIS3D Disp_No 
         Height          =   405
         Left            =   30
         Top             =   450
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   714
         Caption         =   "000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
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
      Begin Cis3D_v60.CIS3D CIS3D19 
         Height          =   405
         Left            =   30
         Top             =   30
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   714
         ForeColor       =   16711680
         Caption         =   "表示連番"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "( F1 )"
         Height          =   225
         Left            =   1380
         TabIndex        =   28
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.PictureBox B1_AreaG 
      Height          =   6765
      Left            =   -2280
      ScaleHeight     =   6705
      ScaleWidth      =   8355
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   8415
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   6585
         Left            =   60
         Top             =   60
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   11615
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
         Begin MSChart20Lib.MSChart MSChart1 
            Height          =   6075
            Left            =   -540
            OleObjectBlob   =   "CKN0420.frx":0000
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   -270
            Width           =   9405
         End
      End
   End
   Begin VB.PictureBox B1_Area0 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H00808000&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   9390
      ScaleHeight     =   1245
      ScaleWidth      =   2085
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   480
      Width           =   2085
      Begin Cis3D_v60.CIS3D Disp_YM 
         Height          =   465
         Left            =   150
         Top             =   570
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   820
         BackColor       =   16777152
         Caption         =   "YYYY年MM月"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16777152
         cBoderColor2    =   16777152
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
         cBoderStyle     =   2
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '透明
         Caption         =   "表示年月"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   420
         TabIndex        =   25
         Top             =   150
         Width           =   1365
      End
      Begin VB.Line Line1 
         X1              =   60
         X2              =   2070
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  '不透明
         Height          =   1065
         Left            =   60
         Shape           =   4  '丸みのある長方形
         Top             =   60
         Width           =   2025
      End
   End
   Begin VB.PictureBox B1_Area3 
      Height          =   2355
      Left            =   90
      ScaleHeight     =   2295
      ScaleWidth      =   8355
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8670
      Width           =   8415
      Begin Cis3D_v60.CIS3D CIS3D30 
         Height          =   765
         Left            =   60
         Top             =   750
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   1349
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
            Height          =   675
            Left            =   1140
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   30
            Width           =   7080
            _Version        =   196608
            _ExtentX        =   12488
            _ExtentY        =   1191
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
            FormatString    =   "№|増減|開 始 日|終 了 日|設定 |便当 |累積 |完 了 日|"
            Rows            =   2
            Cols            =   9
            BackColor       =   16777215
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   0
            ScrollBars      =   2
            RowHeightMin    =   300
            GridLines       =   3
            SelectionMode   =   1
            AllowBigSelection=   0   'False
            BackColorAlternate=   16777215
         End
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   765
            Index           =   6
            Left            =   -30
            Top             =   0
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   1349
            BackColor       =   0
            ForeColor       =   16777215
            Caption         =   "増減設定"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
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
      End
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   705
         Left            =   60
         Top             =   30
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   1244
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
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   705
            Index           =   5
            Left            =   0
            Top             =   0
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   1244
            BackColor       =   0
            ForeColor       =   16777215
            Caption         =   "内示"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
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
         Begin Cis3D_v60.CIS3D CIS3D20 
            Height          =   330
            Left            =   1200
            Top             =   30
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   582
            Caption         =   "内示数"
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
         Begin Cis3D_v60.CIS3D B1lb_NSu 
            Height          =   330
            Left            =   1200
            Top             =   360
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   582
            BackColor       =   14737632
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
         Begin Cis3D_v60.CIS3D CIS3D23 
            Height          =   330
            Left            =   2850
            Top             =   30
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   582
            Caption         =   "日当枚数"
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
         Begin Cis3D_v60.CIS3D B1lb_DaySu 
            Height          =   330
            Left            =   2850
            Top             =   360
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   582
            BackColor       =   14737632
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
         Begin Cis3D_v60.CIS3D CIS3D25 
            Height          =   330
            Left            =   4650
            Top             =   30
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   582
            Caption         =   "回転枚数"
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
         Begin Cis3D_v60.CIS3D B1lb_KSu 
            Height          =   330
            Left            =   4650
            Top             =   360
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   582
            BackColor       =   14737632
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
         Begin Cis3D_v60.CIS3D CIS3D27 
            Height          =   330
            Left            =   6450
            Top             =   30
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   582
            Caption         =   "前月差"
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
         Begin Cis3D_v60.CIS3D B1lb_SSu 
            Height          =   330
            Left            =   6450
            Top             =   360
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   582
            BackColor       =   14737632
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
      End
      Begin Cis3D_v60.CIS3D CIS3D31 
         Height          =   765
         Left            =   60
         Top             =   1530
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   1349
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
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   735
            Index           =   7
            Left            =   -30
            Top             =   0
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   1296
            BackColor       =   0
            ForeColor       =   16777215
            Caption         =   " 平準化 設 定"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
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
            cPositionY      =   -60
         End
         Begin vsFlexLib.vsFlexArray vsGrid2 
            Height          =   690
            Left            =   1110
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   30
            Width           =   7110
            _Version        =   196608
            _ExtentX        =   12541
            _ExtentY        =   1217
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
            FormatString    =   "№|設定 |開 始 日   |終 了 日   |便当枚数|日当枚数|月当枚数|"
            Rows            =   2
            Cols            =   8
            BackColor       =   16777215
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   0
            ScrollBars      =   2
            RowHeightMin    =   300
            GridLines       =   3
            SelectionMode   =   1
            AllowBigSelection=   0   'False
            BackColorAlternate=   16777215
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
      Height          =   1335
      Left            =   11580
      ScaleHeight     =   1275
      ScaleWidth      =   3570
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   420
      Width           =   3630
      Begin CisBtn_60.CisBtn PB_Chg 
         Height          =   585
         Left            =   930
         Top             =   660
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
         BtnCaption      =   "切替"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   8388736
         BtnShadowColor  =   16777215
         PFCaption       =   "( F5 )"
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
         PFBackColor     =   8388736
      End
      Begin CisBtn_60.CisBtn PB_CAN 
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
         Left            =   2670
         Top             =   660
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
         Left            =   2670
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
         Left            =   60
         Top             =   660
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
         Left            =   930
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
         Left            =   60
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
      Height          =   1290
      Left            =   105
      ScaleHeight     =   1230
      ScaleWidth      =   9135
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   450
      Width           =   9195
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   1065
         Left            =   8550
         Top             =   75
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   1879
         ForeColor       =   16711680
         Caption         =   "受"
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
         cPositionY      =   90
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   375
            Left            =   60
            TabIndex        =   3
            Top             =   465
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
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
      Begin Cis3D_v60.CIS3D BLB_HINBN 
         Height          =   1065
         Left            =   1515
         Top             =   75
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   1879
         ForeColor       =   16711680
         Caption         =   "品   番"
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
         cPositionY      =   90
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   60
            TabIndex        =   1
            Top             =   480
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   661
            cWildeCard      =   -1  'True
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
         Height          =   1065
         Left            =   4815
         Top             =   75
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1879
         ForeColor       =   16711680
         Caption         =   "仕  入  先"
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
         cPositionY      =   90
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   60
            TabIndex        =   2
            Top             =   480
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   661
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
            Top             =   480
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXX"
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
      Begin Cis3D_v60.CIS3D CIS3D38 
         Height          =   1065
         Left            =   90
         Top             =   75
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1879
         ForeColor       =   16711680
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
         cPositionY      =   90
         Begin CisText_V60.CisText iH1_Seban 
            Height          =   360
            Left            =   165
            TabIndex        =   0
            Top             =   495
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   635
            cWildeCard      =   -1  'True
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
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   6780
      Left            =   90
      ScaleHeight     =   6720
      ScaleWidth      =   8370
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1860
      Width           =   8430
      Begin vsFlexLib.vsFlexArray vsGrid 
         Height          =   6555
         Left            =   75
         TabIndex        =   4
         Top             =   75
         Width           =   8190
         _Version        =   196608
         _ExtentX        =   14446
         _ExtentY        =   11562
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   "日付    | |１日      |２日      |３日      |４日       |５日    |６日    |７日      |"
         Rows            =   36
         Cols            =   11
         BackColor       =   16777152
         BackColorFixed  =   16777152
         BackColorBkg    =   12632256
         TextStyle       =   4
         TextStyleFixed  =   3
         FocusRect       =   2
         HighLight       =   0
         ScrollBars      =   0
         RowHeightMin    =   200
         FillStyle       =   1
         GridColor       =   14737632
         SelectionMode   =   2
         BackColorAlternate=   16777152
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
      IMEMode         =   3  'ｵﾌ固定
      Left            =   960
      TabIndex        =   11
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
   Begin Cis3D_v60.CIS3D UC_3D11 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 かんばん流動管理問合せ[ｶﾚﾝﾀﾞｰ] 】"
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
         Left            =   14430
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
         Left            =   11910
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
   Begin VB.PictureBox B1_Area2 
      BackColor       =   &H00E0E0E0&
      Height          =   9180
      Left            =   8580
      ScaleHeight     =   9120
      ScaleWidth      =   6585
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1860
      Width           =   6645
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   6555
         Left            =   2925
         Top             =   2490
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   11562
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
         Begin VB.OptionButton B1Ch_DispDay 
            BackColor       =   &H00C0C0C0&
            Caption         =   "回　収　日"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   180
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   8
            Top             =   660
            Width           =   1455
         End
         Begin VB.OptionButton B1Ch_DispDay 
            BackColor       =   &H00C0C0C0&
            Caption         =   "無効指示日"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   1830
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   10
            Top             =   690
            Width           =   1425
         End
         Begin VB.OptionButton B1Ch_DispDay 
            BackColor       =   &H00C0C0C0&
            Caption         =   "有効指示日"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   1830
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   9
            Top             =   330
            Width           =   1455
         End
         Begin VB.OptionButton B1Ch_DispDay 
            BackColor       =   &H00C0C0C0&
            Caption         =   "最終流動日"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   180
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   7
            Top             =   330
            Width           =   1455
         End
         Begin vsFlexLib.vsFlexArray vsGrid3 
            Height          =   5460
            Left            =   90
            TabIndex        =   6
            Top             =   1020
            Width           =   3390
            _Version        =   196608
            _ExtentX        =   5980
            _ExtentY        =   9631
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
            FormatString    =   "№ |か区分|日付|設定|有効指示|無効指示|回収日|"
            Rows            =   17
            Cols            =   9
            FixedCols       =   0
            BackColor       =   12648447
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   320
            GridLines       =   2
            SelectionMode   =   1
            BackColorAlternate=   12648447
         End
         Begin Cis3D_v60.CIS3D CIS3D21 
            Height          =   315
            Left            =   0
            Top             =   0
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            BackColor       =   0
            ForeColor       =   65535
            Caption         =   "【かんばん連番 情報】"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
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
      End
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   2625
         Left            =   60
         Top             =   6420
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   4630
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
         Begin vsFlexLib.vsFlexArray vsGrid4 
            Height          =   1680
            Left            =   120
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   900
            Width           =   2640
            _Version        =   196608
            _ExtentX        =   4657
            _ExtentY        =   2963
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
            FormatString    =   "№|納入日  |便|枚数|||"
            Rows            =   5
            Cols            =   7
            BackColor       =   12640511
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   0
            RowHeightMin    =   320
            GridLines       =   2
            SelectionMode   =   1
            BackColorAlternate=   12640511
         End
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   285
            Left            =   0
            Top             =   0
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   503
            BackColor       =   0
            ForeColor       =   65535
            Caption         =   "【発注情報】"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
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
         Begin Cis3D_v60.CIS3D CIS3D13 
            Height          =   300
            Left            =   120
            Top             =   300
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   529
            Caption         =   "発注枚数合計"
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
         Begin Cis3D_v60.CIS3D CIS3D28 
            Height          =   300
            Left            =   120
            Top             =   600
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   529
            Caption         =   "発注予定枚数"
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
         Begin Cis3D_v60.CIS3D B1lb_HCSu 
            Height          =   300
            Left            =   1650
            Top             =   300
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   14737632
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
            cPositionX      =   -100
         End
         Begin Cis3D_v60.CIS3D B1lb_HYSu 
            Height          =   300
            Left            =   1650
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   14737632
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
            cPositionX      =   -100
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   2415
         Left            =   60
         Top             =   60
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   4260
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
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   330
            Left            =   450
            Top             =   30
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   582
            Caption         =   "サイクル"
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
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   2385
            Index           =   1
            Left            =   15
            Top             =   0
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   4207
            BackColor       =   0
            ForeColor       =   65535
            Caption         =   "※かんばん情報※"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cPositionX      =   110
            cPositionY      =   200
         End
         Begin Cis3D_v60.CIS3D B_Syuyo 
            Height          =   330
            Index           =   1
            Left            =   780
            Top             =   720
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
         Begin Cis3D_v60.CIS3D CIS3D8 
            Height          =   330
            Left            =   1740
            Top             =   30
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D CIS3D12 
            Height          =   330
            Left            =   3060
            Top             =   30
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   582
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
            Height          =   330
            Index           =   1
            Left            =   3630
            Top             =   30
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
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
            Height          =   330
            Index           =   2
            Left            =   4290
            Top             =   30
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
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
            Height          =   330
            Index           =   3
            Left            =   4950
            Top             =   30
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
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
            Height          =   330
            Index           =   0
            Left            =   5610
            Top             =   30
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B1lb_Cycle 
            Height          =   330
            Left            =   450
            Top             =   360
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   582
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
            Height          =   330
            Left            =   1740
            Top             =   360
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B1lb_HMaisu 
            Height          =   330
            Index           =   0
            Left            =   3060
            Top             =   360
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   582
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
            Height          =   330
            Index           =   0
            Left            =   3630
            Top             =   360
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
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
            Height          =   330
            Index           =   0
            Left            =   4290
            Top             =   360
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
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
            Height          =   330
            Index           =   0
            Left            =   4950
            Top             =   360
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
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
            Height          =   330
            Index           =   0
            Left            =   5610
            Top             =   360
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B_Syuyo 
            Height          =   330
            Index           =   2
            Left            =   780
            Top             =   1380
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
         Begin Cis3D_v60.CIS3D B1lb_Syuyo 
            Height          =   330
            Index           =   2
            Left            =   780
            Top             =   1710
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   330
            Index           =   0
            Left            =   450
            Top             =   2040
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B1lb_SYmd 
            Height          =   330
            Left            =   3060
            Top             =   2040
            Width           =   3330
            _ExtentX        =   5874
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B_Hkbn 
            Height          =   330
            Index           =   4
            Left            =   3630
            Top             =   720
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B_Hkbn 
            Height          =   330
            Index           =   5
            Left            =   4290
            Top             =   720
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B_Hkbn 
            Height          =   330
            Index           =   6
            Left            =   4950
            Top             =   720
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B_Hkbn 
            Height          =   330
            Index           =   7
            Left            =   5610
            Top             =   720
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B1lb_YSu 
            Height          =   330
            Index           =   1
            Left            =   3630
            Top             =   1050
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
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
            Height          =   330
            Index           =   1
            Left            =   4290
            Top             =   1050
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
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
            Height          =   330
            Index           =   1
            Left            =   4950
            Top             =   1050
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
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
            Height          =   330
            Index           =   1
            Left            =   5610
            Top             =   1050
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B_Hkbn 
            Height          =   330
            Index           =   8
            Left            =   3630
            Top             =   1380
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B_Hkbn 
            Height          =   330
            Index           =   9
            Left            =   4290
            Top             =   1380
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B_Hkbn 
            Height          =   330
            Index           =   10
            Left            =   4950
            Top             =   1380
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B_Hkbn 
            Height          =   330
            Index           =   11
            Left            =   5610
            Top             =   1380
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B1lb_YSu 
            Height          =   330
            Index           =   2
            Left            =   3630
            Top             =   1710
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
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
            Height          =   330
            Index           =   2
            Left            =   4290
            Top             =   1710
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
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
            Height          =   330
            Index           =   2
            Left            =   4950
            Top             =   1710
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
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
            Height          =   330
            Index           =   2
            Left            =   5610
            Top             =   1710
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B1lb_Syuyo 
            Height          =   330
            Index           =   1
            Left            =   780
            Top             =   1050
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   330
            Left            =   3060
            Top             =   720
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B1lb_HMaisu 
            Height          =   330
            Index           =   1
            Left            =   3060
            Top             =   1050
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D CIS3D16 
            Height          =   330
            Left            =   3060
            Top             =   1380
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B1lb_HMaisu 
            Height          =   330
            Index           =   2
            Left            =   3060
            Top             =   1710
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B_Syuyo 
            Height          =   330
            Index           =   0
            Left            =   1740
            Top             =   720
            Width           =   1320
            _ExtentX        =   2328
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
         Begin Cis3D_v60.CIS3D B1lb_Youki 
            Height          =   330
            Index           =   1
            Left            =   1740
            Top             =   1050
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
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
         Begin Cis3D_v60.CIS3D B_Syuyo 
            Height          =   330
            Index           =   3
            Left            =   1740
            Top             =   1380
            Width           =   1320
            _ExtentX        =   2328
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
         Begin Cis3D_v60.CIS3D B1lb_Youki 
            Height          =   330
            Index           =   2
            Left            =   1740
            Top             =   1710
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
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
            Height          =   705
            Index           =   3
            Left            =   450
            Top             =   690
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   1244
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
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cPositionX      =   60
            cPositionY      =   200
         End
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   645
            Index           =   4
            Left            =   450
            Top             =   1380
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   1138
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
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cPositionX      =   60
            cPositionY      =   200
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   3915
         Left            =   60
         Top             =   2490
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   6906
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
         Begin vsFlexLib.vsFlexArray vsGrid5 
            Height          =   3240
            Left            =   120
            TabIndex        =   5
            Top             =   630
            Width           =   2640
            _Version        =   196608
            _ExtentX        =   4657
            _ExtentY        =   5715
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
            FormatString    =   "連番|納入日  |便|設定|"
            Rows            =   10
            Cols            =   6
            FixedCols       =   0
            BackColor       =   12648447
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   320
            GridLines       =   2
            SelectionMode   =   1
            BackColorAlternate=   12648447
         End
         Begin Cis3D_v60.CIS3D CIS3D15 
            Height          =   315
            Left            =   0
            Top             =   0
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            BackColor       =   0
            ForeColor       =   65535
            Caption         =   "【流動かんばん】"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
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
         Begin Cis3D_v60.CIS3D CIS3D17 
            Height          =   300
            Left            =   120
            Top             =   330
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   529
            Caption         =   "流動日"
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
         Begin Cis3D_v60.CIS3D B1lb_RYmd 
            Height          =   300
            Left            =   1170
            Top             =   330
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            BackColor       =   14737632
            Caption         =   "YYYY/MM/DD"
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
   End
   Begin Cis3D_v60.CIS3D DispSuryo 
      Height          =   600
      Left            =   13320
      Top             =   2550
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1058
      BackColor       =   16711680
      ForeColor       =   16777215
      Caption         =   "数 量"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   21.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      cAlingnment     =   7
   End
   Begin VB.Label ColorToday 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFC0C0&
      Caption         =   "当日色"
      Height          =   330
      Left            =   10125
      TabIndex        =   19
      Top             =   1500
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label ColorSata 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FF8080&
      Caption         =   "土曜色"
      Height          =   330
      Left            =   8955
      TabIndex        =   18
      Top             =   1500
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label ColorSunday 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0C0FF&
      Caption         =   "日曜色"
      Height          =   315
      Left            =   7800
      TabIndex        =   17
      Top             =   1500
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label DayColor 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00E0E0E0&
      Caption         =   "対象外色"
      Height          =   300
      Left            =   6600
      TabIndex        =   16
      Top             =   1500
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Menu Pop_Menu 
      Caption         =   "設定"
      Visible         =   0   'False
      Begin VB.Menu Pop_Yuko 
         Caption         =   "有効"
      End
      Begin VB.Menu Pop_Muko 
         Caption         =   "無効"
      End
      Begin VB.Menu Pop_Clear 
         Caption         =   "取消"
      End
   End
End
Attribute VB_Name = "CKN0420"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   かんばん流動管理問合せ【カレンダー】
'**       フォームID    :   CKN0420
'**       処理概要      :
'**
'**       作  成  日    :   2005/06/14  By CIS
'**       変  更  日    :   2006/06/13  By CIS 増減完了日の表示を修正
'**       変  更  日    :   2006/12/20  By CIS 版数表示（かんばん流動管理テーブル・かんばん流動履歴テーブル）
'**       変  更  日    :   2008/04/11  By CIS 背番号８桁対応
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    Dim CisVsGrid1              As New CisVsGrid3
    Dim CisVsGrid2              As New CisVsGrid3
    Dim CisVsGrid3              As New CisVsGrid3
    Dim CisVsGrid4              As New CisVsGrid3
    Dim CisVsGrid5              As New CisVsGrid3
    
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    
    Dim mWeekStart              As Integer
    Dim mSaturdayColor          As Long
    Dim mSundayColor            As Long

    Dim mDayRow(31)             As Integer
    Dim mDayCol(31)             As Integer

    Dim mSuryo(5, 31)           As Long
    Dim mMaisu(5, 31)           As Long
    Dim mZaiko(31)              As Long
    Dim mNyuka(31)              As Long
    Dim mZaikoNew               As Long
    
    Dim mSuryoK(31)             As Long
    Dim mSuryoA(31)             As Long
    
    Dim mSetType                As Integer
    Dim mCellSet                As Boolean
    Dim mSetDisp                As Integer
    
    Dim wHCSu                   As Long
    Dim mGDansu                 As Integer
    Dim mCurrency               As Currency

    Dim mDispRow                As Long
    Dim mDispCol                As Long
    
    Dim mDispYM                 As String
    Dim mMaisuMax               As Long
    Dim mMaisuMaxIni            As Long

'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》             +
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
    
    B1_AreaG.Move B1_Area1.Left, B1_Area1.Top
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
    
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
    
    mDispYM = Format(Date, "YYYYMM")
    Disp_YM = Mid(mDispYM, 1, 4) & "年" & Mid(mDispYM, 5, 2) & "月"
    mMaisuMaxIni = 10
    ' 曜日開始を求める
    gSL_Select = "select 数字1 From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = '曜日開始管理'"
    gSL_Select = gSL_Select & "   And 値 = '1'"
    If Not SYMRead(gSL_Select, 1) Then
       SYM.数字1 = 0
    End If
    
    mWeekStart = SYM.数字1 + 1

    mSaturdayColor = &HFFC0C0
    mSundayColor = &HC0C0FF
    
    mGDansu = 4
    Call GridInit
'   かんばん増減管理
    With CisVsGrid1
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 300
        .PatanMax = 1
        .SelectCol = 10
        .InitGet = False
        .Init
    End With
    Call GridInit1
'   かんばん平準化管理
    With CisVsGrid2
        Set .GridObj = vsGrid2
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 300
        .PatanMax = 1
        .SelectCol = 8
        .InitGet = False
        .Init
    End With
    Call GridInit2
'   無効･滞留・未流動かんばん
    With CisVsGrid3
        Set .GridObj = vsGrid3
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 4
        .SelectCol = 2
        .InitGet = False
        .Init
    End With
    Call GridInit3

'   発注情報
    With CisVsGrid4
        Set .GridObj = vsGrid4
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 10
        .InitGet = False
        .Init
    End With
    Call GridInit4
'   流動かんばん
    With CisVsGrid5
        Set .GridObj = vsGrid5
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 10
        .InitGet = False
        .Init
    End With
    Call GridInit5

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
    
           Case vbKeyF1:        If Back_No.Visible Then Call Back_No_Click  '【解除】
           Case vbKeyF5:        If PB_Chg.Visible Then Call PB_Chg_Click    '【切替】
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyPageUp:    If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
'                If ActiveControl.Name = "vsGrid3" Then
'                   Exit Function
'                Else
'                   If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
'                End If
           Case vbKeyPageDown:  If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
'                If ActiveControl.Name = "vsGrid3" Then
'                   Exit Function
'                Else
'                   If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
'                End If
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_Can_Click    '【取消】
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                      '【入力】
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
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_Can_Click()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "H1" Then
        Call HeadBodyClear("H")
        iH1_Seban.SetFocus
        GoTo PB_CAN_Ed
    End If
    
    If ProcHB = "B1" Then
        ProcHB = "H1"
        Back_DispNo.Visible = False
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        Call GridInit
        iH1_Seban.SetFocus
        GoTo PB_CAN_Ed
    End If
    
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    If ProcHB = "H1" Then
        Call CisFrm.MousePT(11)

        Call SqlSelectSet("Back")
        Call DBInput("Back")

        On Error Resume Next
        iH1_Seban.SetFocus
        On Error GoTo 0

        Call CisFrm.MousePT(1)
    End If
    If ProcHB = "B1" Then
       Call CisFrm.MousePT(11)
       With CisFun
            .Date_RsFormat = fYM
            .Date_Add mDispYM, M, -1
            mDispYM = .Date_Result
            Disp_YM = Mid(mDispYM, 1, 4) & "年" & Mid(mDispYM, 5, 2) & "月"
            Call BodyRenbanDisp("R")
       End With
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

        On Error Resume Next
        iH1_Seban.SetFocus
        On Error GoTo 0

        Call CisFrm.MousePT(1)
    End If
    If ProcHB = "B1" Then
       Call CisFrm.MousePT(11)
       With CisFun
            .Date_RsFormat = fYM
            .Date_Add mDispYM, M, 1
            mDispYM = .Date_Result
            Disp_YM = Mid(mDispYM, 1, 4) & "年" & Mid(mDispYM, 5, 2) & "月"
            Call BodyRenbanDisp("R")
       End With
       Call CisFrm.MousePT(1)
    End If
End Sub
'+----------------------------+
'+    切　替  キ ー(F5)       +
'+----------------------------+
Private Sub PB_Chg_Click()
    
    If B1_AreaG.Visible Then
       B1_AreaG.Visible = False
       B1_Area1.Visible = True
    Else
       B1_AreaG.Visible = True
       B1_Area1.Visible = False
    End If
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_SSTorcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_UTorcd(0): Exit Sub
    If PB_Look.Tag = "iH1_Seban" Then Call Look_Hinbn(1): Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn(0): Exit Sub
End Sub
'+-----------------------------+
'+       取引先検索             +
'+-----------------------------+
Private Sub Look_SSTorcd()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 8
    
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
Private Sub Look_UTorcd(KB As Integer)
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_Torcd = iH1_Torcd
    RV_TorKb = 8
    RV_Kojyo = ""
    
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
'+       品番検索              +
'+-----------------------------+
Private Sub Look_Hinbn(KB As Integer)
    RV_Left = 0
    RV_Top = 0
        
    RVI_Hinmoku = 1 ' 発注
    RVI_SK = 0
    RV_Call = ""
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iH1_Seban = RV_Seban
        iH1_Hinbn = RV_Hinbn
        iH1_Torcd = RV_TorcdK
        iH1_Ukeir = RV_Ukeir
        Call iH1_Torcd_LostFocus
    End If
End Sub
'+-----------------------------+
'+  表示年月
'+-----------------------------+
Private Sub Disp_YM_Click()
    If ProcHB <> "B1" Then Exit Sub
    
    If mDispYM = Format(Date, "YYYYMM") Then Exit Sub
    mDispYM = Format(Date, "YYYYMM")
    Call BodyRenbanDisp("R")
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
    If TorNmGet(iH1_Torcd, 8) Then
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
'+  B1Ch_DispDay
'+-----------------------------+
Private Sub B1Ch_DispDay_Click(Index As Integer)
    
    For Each gObj In B1Ch_DispDay
        gObj.BackColor = &HC0C0C0
    Next gObj
    B1Ch_DispDay(Index).BackColor = &HFFFFFF
    
    CisVsGrid3.LayOut (Index)
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
    
    H1Mode = False:     H1Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "B1":      B1Mode = True:  B1Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B1_Area2.Enabled = B1Mode
    B1_Area2.BackColor = B1Color
    B1_Area3.Enabled = B1Mode
    B1_Area3.BackColor = B1Color
    B1_AreaG.BackColor = B1Color
    
    If Pro = "B1" Then
       PB_NEXT.BtnCaption = "次月"
       PB_BACK.BtnCaption = "前月"
    Else
       PB_NEXT.BtnCaption = "後退"
       PB_BACK.BtnCaption = "前進"
       If B1_AreaG.Visible Then
          B1_AreaG.Visible = False
          B1_Area1.Visible = True
       End If
    End If
    PB_Look.Visible = H1Mode
    PB_Chg.Visible = B1Mode
    
'   ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
End Function
'**************************************
'*      Head/Body 初期化処理           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String
    
    If ClrType = "ALL" Then wbc_Enm = HeadBody & "*" Else wbc_Enm = Trim(ClrType)
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then .Text = "": .Enabled = True
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
        End With
    Next gObj
    If HeadBody = "B" Then
        Call CisVsGrid1.Clear
        Call CisVsGrid2.Clear
        Call CisVsGrid3.Clear
        Call CisVsGrid4.Clear
        Call CisVsGrid5.Clear
        mDispYM = Format(Date, "YYYYMM")
        Disp_YM = Mid(mDispYM, 1, 4) & "年" & Mid(mDispYM, 5, 2) & "月"
        For Each gObj In B1Ch_DispDay
            gObj.Value = False
            gObj.BackColor = &HC0C0C0
        Next gObj
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
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        
        If Not DBPut Then GoTo ReturnPress_Ed
        
        Call PB_Can_Click
        GoTo ReturnPress_Ed
    End If

ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Dim SetItem     As String
    Dim SetDayStr   As String
    Dim YSuryo      As Long
    Dim YYSuryo     As Long
    Dim wWhere      As String
    Dim wJEndDay    As Long
    
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
       iH1_Seban.SetFocus
       Exit Function
    End If

    gSL_Select = "select * from 品番取引先マスタ"
    gSL_Select = gSL_Select & " where 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   and 取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   and 背番号 = '" & RTrim(iH1_Seban) & "'"
    If Not HTMRead(gSL_Select, 1, 1) Then
        gSL_Select = ""
        gSL_Select = gSL_Select & "Select"
        gSL_Select = gSL_Select & "  htm.背番号"
        gSL_Select = gSL_Select & ", htm.取引先"
        gSL_Select = gSL_Select & ", htm.受入"
        gSL_Select = gSL_Select & ", htm.品番"
        gSL_Select = gSL_Select & ", hnm.表示品番"
        gSL_Select = gSL_Select & ", hnm.社内背番"
        gSL_Select = gSL_Select & ", case when htm.最終発行連番 >= isnull(krt.発行NO, 0)"
        gSL_Select = gSL_Select & "       then htm.最終発行連番"
        gSL_Select = gSL_Select & "       else krt.発行NO"
        gSL_Select = gSL_Select & "  end 最終発行連番"
        gSL_Select = gSL_Select & ", rtrim(trm.略称) 取引先名"
        gSL_Select = gSL_Select & ", htm.手配区分"
        gSL_Select = gSL_Select & " from 品番取引先マスタ htm"
        gSL_Select = gSL_Select & "      left outer join 品番マスタ hnm"
        gSL_Select = gSL_Select & "        on hnm.品番 = htm.品番"
        gSL_Select = gSL_Select & "      left outer join 取引先マスタ trm"
        gSL_Select = gSL_Select & "        on trm.取引先CD = htm.取引先"
        gSL_Select = gSL_Select & "       and trm.取引先区分 = 1"
        gSL_Select = gSL_Select & "      left outer join"
        gSL_Select = gSL_Select & " (select 取引先"
        gSL_Select = gSL_Select & "       , 受入"
        gSL_Select = gSL_Select & "       , 品番"
        gSL_Select = gSL_Select & "       , max(発行NO) 発行NO"
        gSL_Select = gSL_Select & "    from かんばん流動管理テーブル"
        gSL_Select = gSL_Select & "   group by 取引先"
        gSL_Select = gSL_Select & "          , 受入"
        gSL_Select = gSL_Select & "          , 品番"
        gSL_Select = gSL_Select & " ) krt"
        gSL_Select = gSL_Select & "    on krt.取引先 = htm.取引先"
        gSL_Select = gSL_Select & "   and krt.受入 = htm.受入"
        gSL_Select = gSL_Select & "   and krt.品番 = htm.品番"
        gSL_Select = gSL_Select & " where htm.背番号 = '" & iH1_Seban & "'"
        gSL_Select = gSL_Select & "   and htm.品目 = 1"
        gSL_Select = gSL_Select & "   and htm.SK区分 = 0"
        gSL_Select = gSL_Select & " order by"
        gSL_Select = gSL_Select & "  htm.品番"
        gSL_Select = gSL_Select & ", htm.取引先"
        gSL_Select = gSL_Select & ", htm.受入"
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

'           品番
            If Trim(iH1_Hinbn) = "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    品番を入力して下さい            "
                    .MB_Button = Error
                    .MBOX
                End With
                iH1_Hinbn.SetFocus
                Exit Function
            End If
'           取引先
            If Trim(iH1_Torcd) = "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    仕入先を入力して下さい            "
                    .MB_Button = Error
                    .MBOX
                End With
                iH1_Torcd.SetFocus
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
               iH1_Torcd.SetFocus
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
                iH1_Ukeir.SetFocus
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
               iH1_Ukeir.SetFocus
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
            iH1_Seban.SetFocus
            Exit Function

        End If
    End If
    
    Call GridInit
    
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
'
    Head1Chk = True

Head1Chk_Ed:
    PNL_Wait.Visible = False
End Function
'****************************
'*      BODY1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
    With vsGrid3
         gCnt2 = 0
         For gCnt1 = 1 To .Rows - 1
             If Trim(.TextMatrix(gCnt1, 3)) <> "" Then
                gCnt2 = 1
                Exit For
             End If
         Next gCnt1
    End With
    
    If gCnt2 = 0 Then
       With CisFun
           .MB_Lines = 4
           .MB_MSG(2) = "    かんばん連番情報の設定がされていません。            "
           .MB_Title = ""
           .MB_Button = Error
           .MBOX
        End With
        Exit Function
    End If
    
    Body1Chk = True
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示     *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    Call CisFrm.MousePT(11)
    
    DBInput = False
    On Error GoTo DBInput_Err:
    
    Call HeadBodyClear("B")
    If InputKey = "INV" Then
       PNL_Wait.Visible = True: DoEvents
    End If
    
    If Not KRTRead(gSL_Select, 1) Then
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
        iH1_Seban = RTrim(KRT.背番号)
        iH1_Hinbn = RTrim(KRT.品番)
        iH1_Torcd = RTrim(KRT.取引先)
        iH1_Ukeir = RTrim(KRT.受入)
        H1lb_Tornm = RTrim(KRT.取引先名)
        Select Case InputKey
            Case "INV": If Not BodySet("ALL") Then GoTo DBInput_Ed
            Case Else:  If Not BodySet("") Then GoTo DBInput_Ed
        End Select
        DBInput = True
    End If

DBInput_Ed:
    PNL_Wait.Visible = False: DoEvents
    Call CisFrm.MousePT(1)
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
    gSL_Select = gSL_Select & ", htm.サイクル"
    gSL_Select = gSL_Select & ", htm.変更日,htm.変更便"
    gSL_Select = gSL_Select & ", htm.収容数1"
    gSL_Select = gSL_Select & ", htm.収容数2"
    gSL_Select = gSL_Select & ", htm.収容器1"
    gSL_Select = gSL_Select & ", htm.収容器2"
    gSL_Select = gSL_Select & ", htm.置場"
    gSL_Select = gSL_Select & ", trm.略称 取引先名"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  かんばん流動管理テーブル krt"
    gSL_Select = gSL_Select & ", 品番取引先マスタ htm"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ trm"
    gSL_Select = gSL_Select & "    on trm.取引先CD = htm.取引先"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 1"
    gSL_Select = gSL_Select & " where htm.取引先 = krt.取引先"
    gSL_Select = gSL_Select & "   and htm.受入 = krt.受入"
    gSL_Select = gSL_Select & "   and htm.品番 = krt.品番"
    gSL_Select = gSL_Select & "   and"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
    'gSL_Select = gSL_Select & " ( htm.背番号 + space(5 - len(RTrim(htm.背番号)))"
    gSL_Select = gSL_Select & " ( htm.背番号 + space(8 - len(RTrim(htm.背番号)))"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
    gSL_Select = gSL_Select & " + htm.品番 + space(25 - len(RTrim(htm.品番)))"
    gSL_Select = gSL_Select & " + htm.取引先 + space(7 - len(htm.取引先))"
    gSL_Select = gSL_Select & " + htm.受入 + space(2 - len(htm.受入))"
    gSL_Select = gSL_Select & " ) "
    Select Case InputKey
        Case "Back": gSL_Select = gSL_Select & "<"
        Case "Next": gSL_Select = gSL_Select & ">"
        Case Else:   gSL_Select = gSL_Select & "="
    End Select
    gSL_Select = gSL_Select & " ( '"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
    'gSL_Select = gSL_Select & iH1_Seban & Space(5 - CisFun.Len2(iH1_Seban))
    gSL_Select = gSL_Select & iH1_Seban & Space(8 - CisFun.Len2(iH1_Seban))
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
    gSL_Select = gSL_Select & iH1_Hinbn & Space(25 - CisFun.Len2(iH1_Hinbn))
    gSL_Select = gSL_Select & iH1_Torcd & Space(7 - CisFun.Len2(iH1_Torcd))
    gSL_Select = gSL_Select & iH1_Ukeir & Space(2 - CisFun.Len2(iH1_Ukeir))
    gSL_Select = gSL_Select & "' )"
    Select Case InputKey
        Case "Back"
            gSL_Select = gSL_Select & " order by"
            gSL_Select = gSL_Select & "  htm.背番号 desc"
            gSL_Select = gSL_Select & ", htm.品番 desc"
            gSL_Select = gSL_Select & ", htm.取引先 desc"
            gSL_Select = gSL_Select & ", htm.受入 desc"
        Case "Next"
            gSL_Select = gSL_Select & " order by"
            gSL_Select = gSL_Select & "  htm.背番号 asc"
            gSL_Select = gSL_Select & ", htm.品番 asc"
            gSL_Select = gSL_Select & ", htm.取引先 asc"
            gSL_Select = gSL_Select & ", htm.受入 asc"
    End Select
'
End Sub
'+---------------------------------------+
'+  カレンダマスタ読み込み
'+---------------------------------------+
Private Function CalenderRead(Optional MsgDisp As Boolean = True) As Boolean
    CalenderRead = False

    gSL_Select = "Select * From カレンダマスタ"
    gSL_Select = gSL_Select & " Where 取引先区分 = 1"
    gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   年 = '" & Format(Mid(mDispYM, 1, 4), "0000") & "'"
    gSL_Select = gSL_Select & " And   月 = '" & Format(Mid(mDispYM, 5, 2), "00") & "'"
    gSL_Select = gSL_Select & " And   受入 = '" & RTrim(iH1_Ukeir) & "'"
    If Not CLMRead(gSL_Select, 1) Then
       gSL_Select = "Select * From カレンダマスタ"
       gSL_Select = gSL_Select & " Where 取引先区分 = 1"
       gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(iH1_Torcd) & "'"
       gSL_Select = gSL_Select & " And   年 = '" & Format(Mid(mDispYM, 1, 4), "0000") & "'"
       gSL_Select = gSL_Select & " And   月 = '" & Format(Mid(mDispYM, 5, 2), "00") & "'"
       gSL_Select = gSL_Select & " And   受入 = ''"
       If Not CLMRead(gSL_Select, 1) Then
          gSL_Select = "Select * From カレンダマスタ"
          gSL_Select = gSL_Select & " Where 取引先区分 = 1"
          gSL_Select = gSL_Select & " And   取引先 = ''"
          gSL_Select = gSL_Select & " And   年 = '" & Format(Mid(mDispYM, 1, 4), "0000") & "'"
          gSL_Select = gSL_Select & " And   月 = '" & Format(Mid(mDispYM, 5, 2), "00") & "'"
          gSL_Select = gSL_Select & " And   受入 = ''"
          If Not CLMRead(gSL_Select, 1) Then
             If MsgDisp Then
                With CisFun
                   .MB_Lines = 3
                   .MB_MSG(2) = "     カレンダマスタ未登録(自社)  "
                   .MB_Title = "ｶﾚﾝﾀﾞﾏｽﾀ確認"
                   .MB_Button = OK
                   .MBOX
                End With
             End If
             Call ItemsClearCLM
             CLM.稼動区分 = String(31, "0")
             CLM.夜勤区分 = String(31, "0")
             With CisFun
                 .Date_RsFormat = fDD
                 .Date_End mDispYM & "01"
                 CLM.末日 = .Val2(.Date_Result)
                 CLM.稼動日数 = CLM.末日
             End With
          End If
        End If
    End If
    
    CalenderRead = True
End Function
'**************************************
'*          BODY 内容ｾｯﾄ処理           *
'**************************************
Private Function BodySet(SetType As String) As Boolean
    Dim wHYm        As String
    
    BodySet = False

    mMaisuMax = 0

    With KRT
'       かんばん情報
        If Trim(.サイクル) <> "" Then
            B1lb_Cycle = CisFun.RSetFld(CisFun.Mid2(.サイクル, 1, 1), 1, "#") _
                 & "-" & CisFun.RSetFld(CisFun.Mid2(.サイクル, 2, 2), 2, "#") _
                 & "-" & CisFun.RSetFld(CisFun.Mid2(.サイクル, 4, 2), 2, "#")
        End If
        B1lb_Syuyo(1) = Format(.収容数1, "#,###")
        B1lb_Youki(1) = .収容器1
        wHYm = ""
        If Trim(.変更日) <> "" Then
            wHYm = Mid(.変更日, 1, 6)
            B1lb_SYmd = Format(.変更日, "0000/00/00") & " " & .変更便
            B1lb_Syuyo(2) = Format(.収容数2, "#,###")
            B1lb_Youki(2) = .収容器2
        End If
        B1lb_SYmd.Tag = wHYm
        B1lb_Okiba = .置場
    End With

    Call RenbanDisp
    
' 内示情報
    Call BodyNaijiDisp
    
    If SetType <> "ALL" Then
       BodySet = True
       Exit Function
    End If
    
    B1Ch_DispDay(0).Value = True
    
'   発注テーブル（入荷テーブル）
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  hct.納入日"
    gSL_Select = gSL_Select & ", hct.納入便"
    gSL_Select = gSL_Select & ", hct.枚数"
    gSL_Select = gSL_Select & " from 発注テーブル hct"
    gSL_Select = gSL_Select & " where hct.仕入先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   and hct.受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   and hct.品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   and not exists"
    gSL_Select = gSL_Select & " ("
    gSL_Select = gSL_Select & " select nyt.発注管理NO"
    gSL_Select = gSL_Select & "   from 入荷テーブル nyt"
    gSL_Select = gSL_Select & "  where nyt.発注管理NO = hct.発注管理NO"
    gSL_Select = gSL_Select & " )"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  hct.納入日"
    gSL_Select = gSL_Select & ", hct.納入便"
    If HCTRead(gSL_Select, , 1) Then
        wHCSu = 0
        Call GridSet4
        B1lb_HCSu = Format(wHCSu, "#,###")
    End If
    Call HCTClose
'   発注かんばん予定テーブル
    gSL_Select = ""
    gSL_Select = gSL_Select & "select sum(hky.枚数) 枚数"
    gSL_Select = gSL_Select & " from 発注かんばん予定テーブル hky"
    gSL_Select = gSL_Select & " where hky.仕入先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   and hky.受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   and hky.品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   and hky.処理区分 <> 2"
    If HKYRead(gSL_Select, , 1) Then
        B1lb_HYSu = Format(HKY.枚数, "#,###")
    End If
    Call HKYClose

' かんばん増減情報
    Call GridSet1
' かんばん平準化情報
    Call GridSet2
    
    
    Call BodyRenbanDisp
    
    BodySet = True
    
End Function
'+-------------------------------------+
'+
'+       Body内容表示(連番関係)
'+
'+-------------------------------------+
Private Sub BodyRenbanDisp(Optional DispType As String = "")
    Dim wRno        As Integer
    wRno = 0
    If DispType = "R" Then
       Disp_YM = Mid(mDispYM, 1, 4) & "年" & Mid(mDispYM, 5, 2) & "月"
       wRno = CisFun.Val2(Disp_No)
       
       Call GridInit
    
       Call BodyNaijiDisp
    End If
'  連番情報
    Call GridSet3
' かんばん流動情報読込み
    If RyodoRead(wRno) Then
    End If
    Call GridValueSet
' 流動かんばん情報表示
    Call GridSet5
End Sub
'+-------------------------------------+
'+
'+       Body内容表示(内示関係)
'+
'+-------------------------------------+
Private Sub BodyNaijiDisp()
    
    B1lb_KSu = "": B1lb_SSu = ""
    B1lb_NSu = ""
    B1lb_DaySu = ""

    
    gSL_Select = "Select * "
    gSL_Select = gSL_Select & "  From   かんばん回転枚数テーブル "
    gSL_Select = gSL_Select & "   Where 年月 = '" & mDispYM & "'"
    gSL_Select = gSL_Select & "     and 仕入先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "     and 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "     and 受入 = '" & RTrim(iH1_Ukeir) & "'"
    If KKTRead(gSL_Select, 1, 1) Then
        B1lb_KSu = Format(KKT.回転枚数, "###")
        B1lb_SSu = Format(KKT.回転枚数 - KKT.前月回転枚数, "###")
    End If
    
    gSL_Select = "select 内示年月,SUM(数量) 数量 from 発注内示テーブル"
    gSL_Select = gSL_Select & " Where 手配先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   and 内示年月 = '" & mDispYM & "'"
    gSL_Select = gSL_Select & "   and 内示年月 = 計画年月"
    gSL_Select = gSL_Select & "  Group by 内示年月 "
    If HNJRead(gSL_Select, 1, 1) Then
       B1lb_NSu = Format(HNJ.数量, "#,###")
       If mDispYM >= B1lb_SYmd.Tag And Trim(B1lb_SYmd.Tag) <> "" Then
          gCnt1 = CisFun.Val2(B1lb_Syuyo(2))
       Else
          gCnt1 = CisFun.Val2(B1lb_Syuyo(1))
       End If
       ' 日量数
       If HNJ.数量 <> 0 Then
          If Not CalenderRead Then Exit Sub
          mCurrency = HNJ.数量 / CLM.稼動日数 / gCnt1
          gCnt2 = CisFun.HasuProc(CDbl(mCurrency), 切捨て)
          If gCnt2 = 0 And mCurrency <> 0 Then gCnt2 = 1
          B1lb_DaySu = gCnt2
       End If
    End If
End Sub
'+-------------------------------------+
'+
'+       かんばん連番情報表示
'+
'+-------------------------------------+
Private Sub RenbanDisp()
    
    For gInt = 0 To 2
        B1lb_HMaisu(0) = ""
        B1lb_YSu(gInt) = ""
        B1lb_MSu(gInt) = ""
        B1lb_TSu(gInt) = ""
        B1lb_MISu(gInt) = ""
    Next gInt
    For gInt = 1 To 2
        If gInt = 2 And Trim(B1lb_Syuyo(2)) = "" Then Exit For
        gSL_Select = "Select "
        gSL_Select = gSL_Select & "  品番"
        gSL_Select = gSL_Select & ", 取引先"
        gSL_Select = gSL_Select & ", 受入"
        gSL_Select = gSL_Select & ",sum(case when 発行区分 = 0 then 1 else 0 end) 未流件数"
        gSL_Select = gSL_Select & ",sum(case when 発行区分 = 1 then 1 else 0 end) 有効件数"
        gSL_Select = gSL_Select & ",sum(case when 発行区分 = 2 then 1 else 0 end) 無効件数"
        gSL_Select = gSL_Select & ",sum(case when 発行区分 = 3 then 1 else 0 end) 滞留件数"
        gSL_Select = gSL_Select & " from"
        gSL_Select = gSL_Select & "  かんばん流動管理テーブル "
        gSL_Select = gSL_Select & " where 取引先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & "   and 収容区分 = '" & Format(gInt, 0) & "'"
        gSL_Select = gSL_Select & " Group by 品番,取引先,受入"
        If KRTRead(gSL_Select, 1) Then
            B1lb_HMaisu(gInt) = Format(KRT.有効件数 + KRT.無効件数 + KRT.滞留件数 + KRT.未流件数, "###")
            B1lb_YSu(gInt) = Format(KRT.有効件数, "###")
            B1lb_MSu(gInt) = Format(KRT.無効件数, "###")
            B1lb_TSu(gInt) = Format(KRT.滞留件数, "###")
            B1lb_MISu(gInt) = Format(KRT.未流件数, "###")
            B1lb_HMaisu(0) = Format(CisFun.Val2(B1lb_HMaisu(0)) + CisFun.Val2(B1lb_HMaisu(gInt)), "###")
            B1lb_YSu(0) = Format(CisFun.Val2(B1lb_YSu(0)) + CisFun.Val2(B1lb_YSu(gInt)), "###")
            B1lb_MSu(0) = Format(CisFun.Val2(B1lb_MSu(0)) + CisFun.Val2(B1lb_MSu(gInt)), "###")
            B1lb_TSu(0) = Format(CisFun.Val2(B1lb_TSu(0)) + CisFun.Val2(B1lb_TSu(gInt)), "###")
            B1lb_MISu(0) = Format(CisFun.Val2(B1lb_MISu(0)) + CisFun.Val2(B1lb_MISu(gInt)), "###")
        End If
    Next gInt

End Sub
'+-------------------------------------+
'+
'+       かんばん流動情報読込み
'+
'+-------------------------------------+
Private Function RyodoRead(Optional ReadNo As Integer) As Boolean
    RyodoRead = False
    
    Erase mMaisu
    Disp_No = ""
    Back_DispNo.Visible = False
    
    gSL_Select = "Select 収容区分,流動日,"
    gSL_Select = gSL_Select & "Count(*) 件数 "
    gSL_Select = gSL_Select & "    From かんばん流動履歴テーブル "
    gSL_Select = gSL_Select & "  Where 処理区分 = ''"
    gSL_Select = gSL_Select & "    and 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "    and 背番号 = '" & RTrim(iH1_Seban) & "'"
    gSL_Select = gSL_Select & "    And SUbstring(流動日,1,6) = '" & mDispYM & "'"
    gSL_Select = gSL_Select & "    And 取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "    and 受入 = '" & RTrim(iH1_Ukeir) & "'"
    If ReadNo <> 0 Then
       gSL_Select = gSL_Select & "    and 連番 = " & ReadNo
       Disp_No = Format(ReadNo, "000")
       Back_DispNo.Visible = True
    End If
    gSL_Select = gSL_Select & "  Group By 収容区分,流動日"
    gSL_Select = gSL_Select & "  Order by 流動日"
    If Not KRRRead(gSL_Select) Then
       Call KRRClose
       Exit Function
    End If
    
    Do Until Not KRR_RDSTS
        gCnt1 = CisFun.Val2(Mid(KRR.流動日, 7, 2))
        
        mMaisu(1, gCnt1) = mMaisu(1, gCnt1) + KRR.件数
        mMaisu(KRR.収容区分 + 1, gCnt1) = mMaisu(KRR.収容区分 + 1, gCnt1) + KRR.件数
        ' 枚数最大値　退避
        If mMaisu(1, gCnt1) > mMaisuMax Then
           mMaisuMax = mMaisu(1, gCnt1)
        End If
   
        Call KRRReadNext
    Loop
    
    Call KRRClose

    RyodoRead = True

End Function

'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理　関　連
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'
'**************************************
'*                                    *
'*       カレンダグリッド初期化       *
'*                                    *
'**************************************
Private Sub GridInit()
    Dim SetColor        As Boolean
    
    mCellSet = True
    
    With vsGrid
        .Redraw = False
        .Rows = 1: .Cols = 1
        .Rows = mGDansu * 6: .Cols = 9
        
        .Row = 0
        For gCnt2 = 0 To .Cols - 1
            .Col = gCnt2
            .ColAlignment(gCnt2) = flexAlignCenterCenter
            If gCnt2 > 0 Then
                If gCnt2 = 1 Then
                   .ColWidth(gCnt2) = 30
                Else
                   .ColWidth(gCnt2) = 1000
                End If
            End If
        Next gCnt2
        
        Call DaySet
        
        Call ChartInit
        
        For gCnt1 = 0 To .Rows - 1
            .Row = gCnt1: .Col = 0
            If gCnt1 Mod mGDansu = 0 Then
                .Text = "日 付"
                For gCnt2 = 0 To .Cols - 1
                    .Col = gCnt2
                Next gCnt2
            Else
                Select Case gCnt1 Mod mGDansu
                    Case 1
                        .Text = "合 計"
                        Set gObj = B1lb_HMaisu(0)
                    Case 2
                        .Text = "収容数①"
                        Set gObj = B1lb_Syuyo(1)
                    Case 3
                        .Text = "収容数②"
                        Set gObj = B1lb_Syuyo(2)
                End Select
                .CellBackColor = gObj.BackColor
            End If

            For gCnt2 = 2 To .Cols - 1
                .Col = gCnt2
                SetColor = True
                If gCnt1 Mod mGDansu = 0 Then
                   SetColor = False
                End If
                If gCnt1 < mGDansu Then
                   If Trim(.TextMatrix(0, gCnt2)) = "" Then
                      .CellBackColor = DayColor.BackColor
                      SetColor = False
                   End If
                End If
                If gCnt1 > mGDansu * 4 Then
                   If Trim(.TextMatrix(mGDansu * 4, gCnt2)) = "" Then
                      .CellBackColor = DayColor.BackColor
                      SetColor = False
                   End If
                End If
                If gCnt1 > mGDansu * 5 Then
                   If Trim(.TextMatrix(mGDansu * 5, gCnt2)) = "" Then
                      .CellBackColor = DayColor.BackColor
                      SetColor = False
                   End If
                End If
                If SetColor Then
                   .CellBackColor = gObj.BackColor
                   .CellAlignment = flexAlignRightCenter
                End If
                
            Next gCnt2
        Next gCnt1
        
        .Row = mDispRow: .Col = mDispCol
        .Redraw = True
    End With
    mCellSet = False

End Sub
'+-------------------------------------------+
'+
'+      カレンダマスタより日付情報をセット
'+
'+-------------------------------------------+
Private Sub DaySet()
    Dim wDayCol         As Integer
    Dim SetDay          As Integer
    Dim EndDay          As Integer
    Dim wKado           As Boolean
    
    wDayCol = 1
    SetDay = 1
    
'       *  カレンダ部日付 セット  *
    Call CalenderRead(False)
    
    EndDay = CLM.末日
    gDate = DateValue(Mid(mDispYM, 1, 4) & "/" & Mid(mDispYM, 5, 2) & "/01")
    gInt = Weekday(gDate, mWeekStart)
    Erase mDayRow
    Erase mDayCol
    
    With vsGrid
        For gCnt1 = 0 To .Rows - 1 Step mGDansu
            .Row = gCnt1
            For gCnt2 = 1 To .Cols - 2
                If gCnt2 >= gInt Then
                   .Col = gCnt2 + 1
                   If SetDay <= EndDay Then
                      gDate = DateValue(Mid(mDispYM, 1, 4) & "/" & Mid(mDispYM, 5, 2) & "/" & SetDay)
                      gCnt3 = Weekday(gDate, mWeekStart)
                      .CellBackColor = H1lb_Tornm.BackColor
                      
                      wKado = False
                      If Mid(CLM.稼動区分, SetDay, 1) = "0" Or _
                         Mid(CLM.夜勤区分, SetDay, 1) = "0" Then
                         wKado = True
                      End If
                      If Not wKado Then
                         .CellBackColor = ColorSunday.BackColor
                      End If
                      
                      If mWeekStart = 2 And gCnt3 = 7 And Not wKado Then
                         .CellBackColor = ColorSunday.BackColor
                      End If
                      If mWeekStart = 2 And gCnt3 = 6 And Not wKado Then
                         .CellBackColor = ColorSata.BackColor
                      End If
                      If mWeekStart = 1 And gCnt3 = 1 And Not wKado Then
                         .CellBackColor = ColorSunday.BackColor
                      End If
                      If mWeekStart = 1 And gCnt3 = 7 And Not wKado Then
                         .CellBackColor = ColorSata.BackColor
                      End If
                      
                      If Date = gDate Then
                         .CellBackColor = ColorToday.BackColor
                      End If
                      If mDispYM = Format(Date, "YYYYMM") And Day(Date) = SetDay Then
                         .CellBackColor = ColorToday.BackColor
                      End If
                      
                      .TextMatrix(gCnt1, gCnt2 + 1) = SetDay & " (" & WeekdayName(gCnt2, True, mWeekStart) & ")"
                      mDayRow(SetDay) = gCnt1
                      mDayCol(SetDay) = gCnt2 + 1
                      If SetDay = 1 Then
                         mDispRow = mDayRow(1) + 1
                         mDispCol = mDayCol(1)
                      End If
                      If mDispYM = Format(Date, "YYYYMM") And Day(Date) = SetDay Then
                         mDispRow = mDayRow(SetDay) + 1
                         mDispCol = mDayCol(SetDay)
                      End If
                      
                      SetDay = SetDay + 1
                      gInt = 0
                   Else
                      .CellBackColor = DayColor.BackColor
                   End If
                Else
                    .CellBackColor = DayColor.BackColor
                End If
            Next gCnt2
        Next gCnt1
       .Row = mDispRow: .Col = mDispCol
    End With
End Sub
'+--------------------------------------+
'+
'+      日付毎に枚数をセット
'+
'+--------------------------------------+
Private Sub GridValueSet()
    Dim TSuryo(5)    As Long
    Dim SSuryo(5)    As Long
    Dim SetSu        As Long
    
    mCellSet = True
    
    mDispRow = 0
    mDispCol = 0
    
    With vsGrid
        .Redraw = False
        mDispRow = mDayRow(1) + 1
        mDispCol = mDayCol(1)
        For gCnt1 = 1 To 31
            If mDayRow(gCnt1) = 0 And mDayCol(gCnt1) = 0 Then Exit For
            gDate = DateValue(Mid(mDispYM, 1, 4) & "/" & Mid(mDispYM, 5, 2) & "/" & gCnt1)
            .Col = mDayCol(gCnt1)
            For gCnt2 = 1 To mGDansu - 1
                If mSetType = 0 Then
                    SetSu = mMaisu(gCnt2, gCnt1)
                    .TextMatrix(mDayRow(gCnt1) + gCnt2, mDayCol(gCnt1)) = _
                                            Format(SetSu, "#,###")
                Else
                    SetSu = mMaisu(gCnt2, gCnt1)
                    .TextMatrix(mDayRow(gCnt1) + gCnt2, mDayCol(gCnt1)) = _
                                            Format(SetSu, "#,###")
                End If
                If Date = gDate Then
                   .Row = mDayRow(gCnt1) + gCnt2
                   .CellBackColor = ColorToday.BackColor
                End If
                If mDispYM = Format(Date, "YYYYMM") And Day(Date) = gCnt1 Then
                   .Row = mDayRow(gCnt1) + gCnt2
                   .CellBackColor = ColorToday.BackColor
                   mDispRow = mDayRow(gCnt1) + 1
                   mDispCol = mDayCol(gCnt1)
                End If
                
            Next gCnt2
        Next gCnt1
       .Row = mDispRow: .Col = mDispCol
       .Redraw = True
    End With
    
    mCellSet = False
    
    ' グラフ内容表示
    Call ChartDisp
    
End Sub
Private Function GetDay() As Integer
    Dim wCnt        As Long
    GetDay = 0
    With vsGrid
        For wCnt = 1 To 31
            If mDayRow(wCnt) = (.Row \ mGDansu) * mGDansu And _
               mDayCol(wCnt) = .Col Then
               GetDay = wCnt
               Exit Function
            End If
        Next wCnt
    End With
End Function
Private Sub vsGrid_Click()
    Dim SetDay      As Integer
    
    SetDay = GetDay
    gStr = Format(SetDay, "00")
    
    Call GridSet5(gStr)
End Sub

Private Sub vsGrid_RowColChange()
    If mCellSet Then Exit Sub
    
    Call vsGrid_Click
End Sub

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
        Call .FixedSet(中中, 中中, あり, 4, 1, "増減")
        Call .FixedSet(中中, 中中, あり, 10, 2, "開始日")       ' 2006/03/14 桁数変更 11->10
        Call .FixedSet(中中, 中中, あり, 10, 3, "終了日")       ' 2006/03/14 桁数変更 11->10
        Call .FixedSet(中中, 右中, あり, 5, 4, "設定")
        Call .FixedSet(中中, 右中, あり, 5, 5, "便当")
        Call .FixedSet(中中, 右中, あり, 5, 6, "累積")
        Call .FixedSet(中中, 中中, あり, 9, 7, "完了日")        ' 2006/03/14 桁数変更 7->9
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 8, "Select")
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
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  kzt.設定内容区分"
    gSL_Select = gSL_Select & ", kzt.増減区分"
    gSL_Select = gSL_Select & ", kzt.開始日"
    gSL_Select = gSL_Select & ", kzt.終了日"
    gSL_Select = gSL_Select & ", kzt.設定枚数"
    gSL_Select = gSL_Select & ", kzt.便当枚数"
    gSL_Select = gSL_Select & ", kzt.累積枚数"
    gSL_Select = gSL_Select & ", kzt.完了日"
    gSL_Select = gSL_Select & " from かんばん増減テーブル kzt"
    gSL_Select = gSL_Select & " where kzt.取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   and kzt.受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   and kzt.品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   and kzt.設定内容区分 = 0"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  kzt.開始日 desc"
    If KZTRead(gSL_Select, , 1) Then

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
            ID = 0
            Do Until Not KZT_RDSTS
                ID = ID + 1
    
                'No
                .TextMatrix(ID, CisVsGrid1.FixedGet(0)) = ID
                '増減区分
                Select Case KZT.増減区分
                    Case 0: .TextMatrix(ID, CisVsGrid1.FixedGet(1)) = "減"
                    Case 1: .TextMatrix(ID, CisVsGrid1.FixedGet(1)) = "増"
                End Select
                '開始日
                If KZT.開始日 <> "" Then
                    .TextMatrix(ID, CisVsGrid1.FixedGet(2)) = Format(KZT.開始日, "0000/00/00")
                End If
                '終了日
                If KZT.終了日 <> "" Then
                    .TextMatrix(ID, CisVsGrid1.FixedGet(3)) = Format(KZT.終了日, "0000/00/00")
                End If
                '設定枚数
                .TextMatrix(ID, CisVsGrid1.FixedGet(4)) = Format(KZT.設定枚数, "#,###")
                '便当枚数
                .TextMatrix(ID, CisVsGrid1.FixedGet(5)) = Format(KZT.便当枚数, "#,###")
                '累積枚数
                .TextMatrix(ID, CisVsGrid1.FixedGet(6)) = Format(KZT.累積枚数, "#,###")
                '完了日
                If KZT.完了日 <> "" Then
        '------------------------------------------------------------ 2006/03/14 Delete
'                    .TextMatrix(ID, CisVsGrid1.FixedGet(7)) = Format(KZT.完了日, "0000/00/00")
        '------------------------------------------------------------ 2006/03/14 Delete End
        '------------------------------------------------------------ 2006/03/14 Insert
                    gStr = Format(KZT.完了日, "0000/00/00")
                    .TextMatrix(ID, CisVsGrid1.FixedGet(7)) = Mid(gStr, 3)
        '------------------------------------------------------------ 2006/03/14 Insert
                End If
    
                Call KZTReadNext(1)
            Loop
             .Row = 1: .Col = 1: .ColSel = .Cols - 1
             .Redraw = True
        End With

        GridSet1 = True
    End If
    Call KZTClose
End Function

'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ - かんばん平準化管理
'+----------------------------------------------------------+
Private Sub GridInit2()
    With CisVsGrid2
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 2, 0, "№")
        Call .FixedSet(中中, 中中, あり, 13, 1, "開 始 日")
        Call .FixedSet(中中, 中中, あり, 13, 2, "終 了 日")
        Call .FixedSet(中中, 右中, あり, 8, 3, "便 当")
        Call .FixedSet(中中, 右中, あり, 8, 4, "日 当")
        Call .FixedSet(中中, 右中, あり, 8, 5, "月 当")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 6, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5")
        .InitDisp
    End With
End Sub
'****************************************************'
'*****      グリッド内容セット - かんばん平準化管理
'****************************************************'
Private Function GridSet2() As Boolean
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

    GridSet2 = False
    Call CisVsGrid2.Clear
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  kjt.設定内容区分"
    gSL_Select = gSL_Select & ", kjt.開始日"
    gSL_Select = gSL_Select & ", kjt.終了日"
    gSL_Select = gSL_Select & ", kjt.便上限枚数"
    gSL_Select = gSL_Select & ", kjt.日上限枚数"
    gSL_Select = gSL_Select & ", kjt.月上限枚数"
    gSL_Select = gSL_Select & " from かんばん平準化テーブル kjt"
    gSL_Select = gSL_Select & " where kjt.取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   and kjt.受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   and kjt.品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   and kjt.設定内容区分 = 0"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  kjt.開始日 desc"
    If KJTRead(gSL_Select, , 1) Then

'       抽出処理 初期ｾｯﾄ
        Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
'       ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
        With vsGrid2
           .Redraw = False
            If Ttl_Kensu + 1 < CisVsGrid2.DispMax Then
                .Rows = CisVsGrid2.DispMax
            Else
                .Rows = Ttl_Kensu + 1
            End If
            ID = 0
            Do Until Not KJT_RDSTS
                ID = ID + 1
    
                'No
                .TextMatrix(ID, CisVsGrid2.FixedGet(0)) = ID
                '開始日
                If KJT.開始日 <> "" Then
                    .TextMatrix(ID, CisVsGrid2.FixedGet(1)) = Format(KJT.開始日, "0000/00/00")
                End If
                '終了日
                If KJT.終了日 <> "" Then
                    .TextMatrix(ID, CisVsGrid2.FixedGet(2)) = Format(KJT.終了日, "0000/00/00")
                End If
                '便当枚数
                .TextMatrix(ID, CisVsGrid2.FixedGet(3)) = Format(KJT.便上限枚数, "#,###")
                '日当枚数
                .TextMatrix(ID, CisVsGrid2.FixedGet(4)) = Format(KJT.日上限枚数, "#,###")
                '月当枚数
                .TextMatrix(ID, CisVsGrid2.FixedGet(5)) = Format(KJT.月上限枚数, "#,###")
    
                Call KJTReadNext(1)
            Loop
            .Row = 1: .Col = 1: .ColSel = .Cols - 1
            .Redraw = True
        End With

        GridSet2 = True
    End If
    Call KJTClose
End Function
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ - 無効･滞留・未流動かんばん
'+----------------------------------------------------------+
Private Sub GridInit3()
    With CisVsGrid3
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 中中, あり, 4, 0, "連番")
        Call .FixedSet(中中, 中中, あり, 4, 1, "か区")
        Call .FixedSet(中中, 中中, あり, 10, 2, "最終流動日")
        Call .FixedSet(中中, 中中, あり, 2, 3, "設")
        Call .FixedSet(中中, 中中, あり, 10, 4, "有効指示日")
        Call .FixedSet(中中, 中中, あり, 10, 5, "無効指示日")
        Call .FixedSet(中中, 中中, あり, 10, 6, "回 収 日")
        Call .FixedSet(中中, 右中, あり, 2, 7, "版")                        '2006/12/20追加
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
       'Call .FixedSet(中中, 左中, なし, 0, 2, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,7,1,2,3")
        Call .PatanSet(1, "0,7,1,4,3")
        Call .PatanSet(2, "0,7,1,5,3")
        Call .PatanSet(3, "0,7,1,6,3")
        .InitDisp
    End With
End Sub
'***********************************************'
'*****      グリッド内容セット - 無効･滞留・未流動かんばん
'***********************************************'
Private Sub GridSet3()
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
    
    gSL_Select = "select "
    gSL_Select = gSL_Select & "  発行NO"
    gSL_Select = gSL_Select & ", 流動日付"
    gSL_Select = gSL_Select & ", 発行区分"
    gSL_Select = gSL_Select & ", 有効指示日"
    gSL_Select = gSL_Select & ", 無効指示日"
    gSL_Select = gSL_Select & ", 回収日"
    gSL_Select = gSL_Select & ", 収容区分"
    gSL_Select = gSL_Select & ", 版数"
    gSL_Select = gSL_Select & " from かんばん流動管理テーブル "
    gSL_Select = gSL_Select & " where 取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  発行NO"
    If Not KRTRead(gSL_Select, , 1) Then
       Call KRTClose
       Exit Sub
    End If
' 抽出処理 初期ｾｯﾄ
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid3
       .Redraw = False
        If Ttl_Kensu + 1 < CisVsGrid3.DispMax Then
            .Rows = CisVsGrid3.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
        ID = 0
        Do Until Not KRT_RDSTS
            ID = ID + 1
    
            With vsGrid3
                .TextMatrix(ID, CisVsGrid3.FixedGet(0)) = Format(KRT.発行NO, "000")
                .Row = ID
                If KRT.収容区分 = 2 Then
                   For gInt = 0 To .Cols - 1
                       .Col = gInt
                       .CellBackColor = B1lb_Syuyo(2).BackColor
                   Next gInt
                End If
                .Col = CisVsGrid3.FixedGet(1): .CellBackColor = B_Color(KRT.発行区分).BackColor
                Select Case KRT.発行区分
                    Case 0: gStr = "未流"
                    Case 1: gStr = "有効"
                    Case 2: gStr = "無効"
                    Case 3: gStr = "滞留"
                End Select
                .TextMatrix(ID, CisVsGrid3.FixedGet(1)) = gStr
                If KRT.発行区分 = 1 Or KRT.発行区分 = 3 Then
                   If Trim(KRT.流動日付) <> "" Then
                      With CisFun
                           .DateE_BefVal = KRT.流動日付
                           .DateEdit
                           gStr = Mid(.DateE_AfVal, 3)
                      End With
                      .TextMatrix(ID, CisVsGrid3.FixedGet(2)) = gStr
                   End If
                End If
                If Trim(KRT.有効指示日) <> "" Then
                   With CisFun
                        .DateE_BefVal = KRT.有効指示日
                        .DateEdit
                        gStr = Mid(.DateE_AfVal, 3)
                   End With
                   .TextMatrix(ID, CisVsGrid3.FixedGet(4)) = gStr
                End If
                If Trim(KRT.無効指示日) <> "" Then
                   With CisFun
                        .DateE_BefVal = KRT.無効指示日
                        .DateEdit
                        gStr = Mid(.DateE_AfVal, 3)
                   End With
                   .TextMatrix(ID, CisVsGrid3.FixedGet(5)) = gStr
                End If
                If Trim(KRT.回収日) <> "" Then
                   With CisFun
                        .DateE_BefVal = KRT.回収日
                        .DateEdit
                        gStr = Mid(.DateE_AfVal, 3)
                   End With
                   .TextMatrix(ID, CisVsGrid3.FixedGet(6)) = gStr
                End If
            '------------------------------------------------------------ 2006/12/20 追加 (Start)
                If KRT.版数 <> 0 Then
                   .TextMatrix(ID, CisVsGrid3.FixedGet(7)) = KRT.版数
                End If
            '------------------------------------------------------------ 2006/12/20 追加 (End)
            End With
    
            Call KRTReadNext(1)
        Loop
        .Row = 1: .Col = 1: .ColSel = .Cols - 1
        .Redraw = True
    End With
End Sub

'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ - 発注情報
'+----------------------------------------------------------+
Private Sub GridInit4()
    With CisVsGrid4
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 2, 0, "№")
        Call .FixedSet(中中, 中中, あり, 8, 1, "納入日")
        Call .FixedSet(中中, 中中, あり, 2, 2, "便")
        Call .FixedSet(中中, 右中, あり, 4, 3, "枚数")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
       'Call .FixedSet(中中, 左中, なし, 0, 4, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3")
        .InitDisp
    End With
End Sub
'***********************************************'
'*****      グリッド内容セット - 発注情報
'***********************************************'
Private Sub GridSet4()
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid4
       .Redraw = False
        If Ttl_Kensu + 1 < CisVsGrid4.DispMax Then
            .Rows = CisVsGrid4.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
        
        ID = 0
        Do Until Not HCT_RDSTS
            ID = ID + 1

            .TextMatrix(ID, CisVsGrid4.FixedGet(0)) = ID
            With CisFun
                .DateE_BefVal = HCT.納入日
                .DateEdit
                gStr = Mid(.DateE_AfVal, 3, 8)
            End With
            .TextMatrix(ID, CisVsGrid4.FixedGet(1)) = gStr
            .TextMatrix(ID, CisVsGrid4.FixedGet(2)) = HCT.納入便
            .TextMatrix(ID, CisVsGrid4.FixedGet(3)) = Format(HCT.枚数, "###")
            wHCSu = wHCSu + HCT.枚数
    
            Call HCTReadNext(1)
        Loop
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ - 流動かんばん
'+----------------------------------------------------------+
Private Sub GridInit5()
    With CisVsGrid5
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 中中, あり, 4, 0, "連番")
        Call .FixedSet(中中, 中中, あり, 8, 1, "納入日")
        Call .FixedSet(中中, 中中, あり, 2, 2, "便")
        Call .FixedSet(中中, 右中, あり, 2, 3, "版")        '2006/12/20追加
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
       'Call .FixedSet(中中, 左中, なし, 0, 4, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,3,1,2")
        .InitDisp
    End With
End Sub
'***********************************************'
'*****      グリッド内容セット - 流動かんばん
'***********************************************'
Private Sub GridSet5(Optional DispDay As String = "")
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
    Dim Disp_Ymd    As String
    
    Call CisVsGrid5.Clear
    B1lb_RYmd = ""
    
    If DispDay = "00" Then Exit Sub
    
    If Trim(DispDay) = "" Then
       If Format(Date, "YYYYMM") <> mDispYM Then
          DispDay = "01"
       Else
          DispDay = Format(Day(Date), "00")
       End If
    End If
    
    Disp_Ymd = mDispYM & RTrim(DispDay)
    
    With CisFun
        .DateE_BefVal = Disp_Ymd
        .DateEdit
        B1lb_RYmd = .DateE_AfVal
    End With
    
    gSL_Select = "Select * "
    gSL_Select = gSL_Select & "    From かんばん流動履歴テーブル "
    gSL_Select = gSL_Select & "  Where 処理区分 = ''"
    gSL_Select = gSL_Select & "    and 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "    and 背番号 = '" & RTrim(iH1_Seban) & "'"
    gSL_Select = gSL_Select & "    and 流動日 = '" & Disp_Ymd & "'"
    gSL_Select = gSL_Select & "    And 取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "    and 受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "  Order by 連番,流動日"
    If Not KRRRead(gSL_Select, , 1) Then
       Call KRRClose(1)
       Exit Sub
    End If
    
' 抽出処理 初期ｾｯﾄ
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid5
       .Redraw = False
        If Ttl_Kensu + 1 < CisVsGrid5.DispMax Then
            .Rows = CisVsGrid5.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
        
        ID = 0
        Do Until Not KRR_RDSTS
            ID = ID + 1
            .TextMatrix(ID, CisVsGrid5.FixedGet(0)) = Format(KRR.連番, "000")
            With CisFun
                .DateE_BefVal = KRR.納入日
                .DateEdit
                gStr = CisFun.Mid2(.DateE_AfVal, 3)
            End With
            .TextMatrix(ID, CisVsGrid5.FixedGet(1)) = gStr
            .TextMatrix(ID, CisVsGrid5.FixedGet(2)) = KRR.納入便
            If KRR.収容区分 = 2 Then
               .Row = ID
               For gInt = 0 To .Cols - 1
                   .Col = gInt
                   .CellBackColor = B1lb_Syuyo(2).BackColor
               Next gInt
            End If
            '------------------------------------------------------------ 2006/12/20 追加 (Start)
            If KRR.版数 <> 0 Then
                .TextMatrix(ID, CisVsGrid5.FixedGet(3)) = KRR.版数
            End If
            '------------------------------------------------------------ 2006/12/20 追加 (End)
    
            Call KRRReadNext(1)
        Loop
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    Call KRRClose(1)

End Sub
'+---------------------------------------+
'+      連番情報表示 & 発行区分変更
'+---------------------------------------+
'+-------------- 連番情報解除
Private Sub Back_No_Click()
' かんばん流動情報読込み
    Call GridInit
    If RyodoRead Then
       Call GridValueSet
    End If
End Sub

'+--------------------------------
'+  かんばん連番 情報 設定
'+--------------------------------
Private Sub VsGrid3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        
       With vsGrid3
            .SetFocus
           Pop_Yuko.Visible = True
           Pop_Muko.Visible = True
           Pop_Clear.Visible = True
           If .Row <= .MouseRow And .RowSel >= .MouseRow And .Row <> .RowSel Then
              If Trim(.TextMatrix(.Row, 0)) = "" Then Exit Sub
              If .Row < 1 Then Exit Sub
           Else
              .Row = .MouseRow
              If Trim(.TextMatrix(.Row, 0)) = "" Then Exit Sub
              If Trim(.TextMatrix(.Row, 1)) = "有効" Then
                 Pop_Yuko.Visible = False
              Else
                 Pop_Yuko.Visible = True
              End If
              If Trim(.TextMatrix(.Row, 1)) = "無効" Then
                 Pop_Muko.Visible = False
              Else
                 Pop_Muko.Visible = True
              End If
                
              If Trim(.TextMatrix(.Row, 3)) = "" Then
                 Pop_Clear.Visible = False
              Else
                 Pop_Clear.Visible = True
              End If
           End If
           
       End With
        
        
        
        PopupMenu Pop_Menu
    End If
End Sub
Private Sub vsGrid3_DblClick()
    Set gObj = vsGrid3
    Call NoDisp
End Sub
Private Sub VsGrid3_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeySpace Then Exit Sub
    Set gObj = vsGrid3
    Call NoDisp
End Sub
Private Sub Pop_Yuko_Click()
    With vsGrid3
        For gCnt1 = .Row To .RowSel
            If Trim(.TextMatrix(gCnt1, 0)) = "" Then Exit For
            If Trim(.TextMatrix(gCnt1, 1)) <> "有効" Then
               .TextMatrix(gCnt1, 3) = "有"
            Else
               .TextMatrix(gCnt1, 3) = ""
            End If
        Next gCnt1
    End With
        
End Sub
Private Sub Pop_Muko_Click()
    With vsGrid3
        For gCnt1 = .Row To .RowSel
            If Trim(.TextMatrix(gCnt1, 0)) = "" Then Exit For
            If Trim(.TextMatrix(gCnt1, 1)) <> "無効" Then
               .TextMatrix(gCnt1, 3) = "無"
            Else
               .TextMatrix(gCnt1, 3) = ""
            End If
        Next gCnt1
    End With
End Sub
Private Sub Pop_Clear_Click()
    With vsGrid3
        For gCnt1 = .Row To .RowSel
            If Trim(.TextMatrix(gCnt1, 0)) = "" Then Exit For
            .TextMatrix(gCnt1, 3) = ""
        Next gCnt1
    End With
End Sub

'+--------------------------------
'+  流動かんばん 設定
'+--------------------------------
Private Sub vsGrid5_DblClick()
    Set gObj = vsGrid5
    Call NoDisp
End Sub
Private Sub VsGrid5_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeySpace Then Exit Sub
    Set gObj = vsGrid5
    Call NoDisp
End Sub

Private Sub NoDisp()
    Dim DispNo  As Integer
    With gObj
        If Trim(.TextMatrix(.Row, 0)) = "" Then Exit Sub
        DispNo = CisFun.Val2(.TextMatrix(.Row, 0))
    End With
' かんばん流動情報読込み
    Call GridInit
    If RyodoRead(DispNo) Then
    End If
    Call GridValueSet
End Sub
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    Dim HKbn  As Integer
    Dim Renban As Integer
    DBPut = False

    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "    更新処理を行います。            "
        .MB_MSG(4) = "     よろしいですか？            "
        .MB_Title = "更新処理"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err

    Call CisDB.DBTran(TransBegin)   ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

' ( 更新処理 )
    With vsGrid3
        For gCnt1 = 1 To .Rows - 1
            If Trim(.TextMatrix(gCnt1, 0)) = "" Then Exit For
            HKbn = 0
            If Trim(.TextMatrix(gCnt1, 3)) = "有" Then
               HKbn = 1
            End If
            If Trim(.TextMatrix(gCnt1, 3)) = "無" Then
               HKbn = 2
            End If
            Renban = CisFun.Val2(.TextMatrix(gCnt1, 0))
            If Renban <> 0 And HKbn <> 0 Then
                With KRT
                    .取引先 = RTrim(iH1_Torcd)
                    .受入 = RTrim(iH1_Ukeir)
                    .品番 = RTrim(iH1_Hinbn)
                    .発行NO = Renban
                    .発行区分 = HKbn
                    Select Case .発行区分
                        Case 1
                            .有効指示日 = Format(Date, "yyyymmdd")
                            .無効指示日 = ""
                        Case 2
                            .有効指示日 = ""
                            .無効指示日 = Format(Date, "yyyymmdd")
                    End Select
                    .回収日 = ""
                    .更新者 = gTanto
            
                    gSL_Select = ""
                    gSL_Select = gSL_Select & "update かんばん流動管理テーブル"
                    gSL_Select = gSL_Select & "   set 発行区分 = " & .発行区分 & ""
                    gSL_Select = gSL_Select & "     , 有効指示日 = '" & RTrim(.有効指示日) & "'"
                    gSL_Select = gSL_Select & "     , 無効指示日 = '" & RTrim(.無効指示日) & "'"
                    gSL_Select = gSL_Select & "     , 回収日 = '" & RTrim(.回収日) & "'"
                    gSL_Select = gSL_Select & "     , 更新日 = getdate()"
                    gSL_Select = gSL_Select & "     , 更新者 = '" & RTrim(.更新者) & "'"
                    gSL_Select = gSL_Select & "     , 更新端末 = substring(host_name(),1,20)"
                    gSL_Select = gSL_Select & " where 取引先 = '" & RTrim(.取引先) & "'"
                    gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(.受入) & "'"
                    gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(.品番) & "'"
                    gSL_Select = gSL_Select & "   and 発行NO = " & .発行NO & ""
                End With
                With CisDB
                    .ConnectNo = 0
                    .SQL = gSL_Select
                    If Not .DBExec Then Exit Function
                End With
            End If
        Next gCnt1
    End With

    CisDB.DBTran (TransCommit)      ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+---------------------------------------+
'+      グラフ処理
'+---------------------------------------+
Private Sub ChartInit()
    Dim wKado       As Boolean
    Dim wDispDay    As String
    
    With MSChart1
        .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = mMaisuMaxIni
        .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 5
        
        .ColumnCount = 2
        .RowCount = CLM.末日
'
        For gLong = 1 To CLM.末日       'グラフ初期化
            .Row = gLong
            gDate = DateValue(Mid(mDispYM, 1, 4) & "/" & Mid(mDispYM, 5, 2) & "/" & gLong)
            gCnt3 = Weekday(gDate, mWeekStart)
            wKado = False
            gStr = ""
            If Mid(CLM.稼動区分, gLong, 1) = "0" Or _
               Mid(CLM.夜勤区分, gLong, 1) = "0" Then
               wKado = True
            Else
               gStr = "1"
            End If
            If mWeekStart = 2 And gCnt3 = 7 And Not wKado Then
               gStr = "1"
            End If
            If mWeekStart = 2 And gCnt3 = 6 And Not wKado Then
               gStr = "2"
            End If
            If mWeekStart = 1 And gCnt3 = 1 And Not wKado Then
               gStr = "1"
            End If
            If mWeekStart = 1 And gCnt3 = 7 And Not wKado Then
               gStr = "2"
            End If
'            If gLong = 1 Then
'               wDispDay = CisFun.RSetFld(Mid(mDispYM, 5, 2), 2, "#0") & "/" & CisFun.RSetFld(gLong, 2, "#0")
'            Else
'               wDispDay = "   " & CisFun.RSetFld(gLong, 2, "#0")
'            End If
            wDispDay = CisFun.RSetFld(gLong, 2, "#0")
            Select Case gStr
                Case "1": .RowLabel = "*" & wDispDay    ' 日曜
                Case "2": .RowLabel = "*" & wDispDay    ' 土曜
                Case Else
                    .RowLabel = wDispDay
            End Select
            
            For gInt = 1 To .ColumnCount
                .Column = gInt
                .Data = ""
            Next gInt
        Next gLong
    End With
End Sub
'+---------------------------------------+
'+      グラフ表示
'+---------------------------------------+
Private Sub ChartDisp()
    Dim wNDauMai    As Long

    wNDauMai = CisFun.Val2(B1lb_DaySu)
    With MSChart1
        If mMaisuMax = 0 Then
           mMaisuMax = mMaisuMaxIni
        End If
        If mMaisuMax < wNDauMai Then
           mMaisuMax = wNDauMai
        End If
        If mMaisuMax Mod 5 <> 0 Then
           mMaisuMax = mMaisuMax + (5 - mMaisuMax Mod 5)
        End If
        
        .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = mMaisuMax

        .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 5
        If wNDauMai <> 0 Then
           .ColumnCount = 2
           .Plot.SeriesCollection.Item(2).Pen.VtColor.Set 255, 0, 0
        Else
           .ColumnCount = 1
        End If
        .RowCount = CLM.末日
'
        For gLong = 1 To CLM.末日
            .Row = gLong
            .Column = 1
            .Data = mMaisu(1, gLong)
            If wNDauMai <> 0 Then
               .Column = 2
               .Data = wNDauMai
            End If
        Next gLong
    End With
End Sub

