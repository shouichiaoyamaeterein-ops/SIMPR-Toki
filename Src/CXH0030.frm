VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{2BB9E22C-2B4F-11D4-806A-00000EA57F9E}#5.0#0"; "CISNFT~1.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CisYMDwB_6.0.ocx"
Begin VB.Form CXH0030 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "かんばん発注指示"
   ClientHeight    =   10050
   ClientLeft      =   90
   ClientTop       =   615
   ClientWidth     =   14760
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
   ScaleHeight     =   10050
   ScaleWidth      =   14760
   WindowState     =   2  '最大化
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   645
      Left            =   630
      ScaleHeight     =   585
      ScaleWidth      =   2265
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   915
      Width           =   2325
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   570
         Left            =   15
         Top             =   0
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   1005
         Caption         =   " 表示件数"
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
         Begin Cis3D_v60.CIS3D B1lb_DispSu 
            Height          =   435
            Left            =   1185
            Top             =   60
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   767
            BackColor       =   12648447
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   15.75
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
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   645
      Left            =   2940
      ScaleHeight     =   585
      ScaleWidth      =   2265
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   915
      Width           =   2325
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   570
         Left            =   15
         Top             =   0
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   1005
         Caption         =   " 選択件数"
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
         Begin Cis3D_v60.CIS3D B1lb_SelSu 
            Height          =   435
            Left            =   1215
            Top             =   60
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   767
            BackColor       =   12648384
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   15.75
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
   End
   Begin Cis3D_v60.CIS3D B1_Area1 
      Height          =   7050
      Left            =   630
      Top             =   2985
      Width           =   13845
      _ExtentX        =   24421
      _ExtentY        =   12435
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
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   6900
         Left            =   75
         Top             =   60
         Width           =   13680
         _ExtentX        =   24130
         _ExtentY        =   12171
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
            DragIcon        =   "CXH0030.frx":0000
            Height          =   6705
            Left            =   105
            TabIndex        =   8
            Top             =   105
            Width           =   13515
            _Version        =   196608
            _ExtentX        =   23839
            _ExtentY        =   11827
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
            FormatString    =   $"CXH0030.frx":030A
            Rows            =   21
            Cols            =   20
            BackColor       =   12648447
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   320
            GridLines       =   2
            SelectionMode   =   1
            Begin Cis3D_v60.CIS3D B2_Area1 
               Height          =   4440
               Left            =   4080
               Top             =   1935
               Visible         =   0   'False
               Width           =   8610
               _ExtentX        =   15187
               _ExtentY        =   7832
               BackColor       =   0
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
               Begin Cis3D_v60.CIS3D CIS3D4 
                  Height          =   480
                  Left            =   105
                  Top             =   90
                  Width           =   8415
                  _ExtentX        =   14843
                  _ExtentY        =   847
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
                  Begin VB.Label B2lb_Torcd 
                     BackStyle       =   0  '透明
                     Caption         =   "XXXXX：ＸＸＸＸＸＸＸＸＸＸＸ"
                     Height          =   225
                     Left            =   4095
                     TabIndex        =   17
                     Top             =   120
                     Width           =   4215
                  End
                  Begin VB.Shape Shape4 
                     BackColor       =   &H00C0FFC0&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   3990
                     Shape           =   4  '丸みのある長方形
                     Top             =   45
                     Width           =   4365
                  End
                  Begin VB.Label Label3 
                     BackStyle       =   0  '透明
                     Caption         =   "仕入先"
                     Height          =   225
                     Left            =   3180
                     TabIndex        =   16
                     Top             =   120
                     Width           =   735
                  End
                  Begin VB.Shape Shape3 
                     BackColor       =   &H00C0FFC0&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   3090
                     Shape           =   4  '丸みのある長方形
                     Top             =   45
                     Width           =   885
                  End
                  Begin VB.Label B2lb_Nymd 
                     BackStyle       =   0  '透明
                     Caption         =   "YYYY/MM/DD-99"
                     Height          =   225
                     Left            =   1425
                     TabIndex        =   15
                     Top             =   120
                     Width           =   1560
                  End
                  Begin VB.Shape Shape2 
                     BackColor       =   &H00C0FFC0&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   1335
                     Shape           =   4  '丸みのある長方形
                     Top             =   45
                     Width           =   1770
                  End
                  Begin VB.Label Label1 
                     BackStyle       =   0  '透明
                     Caption         =   "納入日-便"
                     Height          =   225
                     Left            =   150
                     TabIndex        =   14
                     Top             =   120
                     Width           =   1080
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H00C0FFC0&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   60
                     Shape           =   4  '丸みのある長方形
                     Top             =   45
                     Width           =   1275
                  End
               End
               Begin vsFlexLib.vsFlexArray VsGrid2 
                  Height          =   3750
                  Left            =   90
                  TabIndex        =   13
                  TabStop         =   0   'False
                  Top             =   570
                  Width           =   8430
                  _Version        =   196608
                  _ExtentX        =   14870
                  _ExtentY        =   6615
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
                  FormatString    =   "No|      品     番      |背番号|収容数|枚数|発注数|サイクル|"
                  Rows            =   11
                  Cols            =   8
                  BackColor       =   16777152
                  BackColorBkg    =   14737632
                  FocusRect       =   0
                  HighLight       =   2
                  ScrollBars      =   2
                  RowHeightMin    =   330
                  SelectionMode   =   1
                  BackColorAlternate=   16777152
               End
            End
            Begin Cis3D_v60.CIS3D Message 
               Height          =   930
               Left            =   3765
               Top             =   1305
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
               Begin Cis3D_v60.CIS3D MessageDisp 
                  Height          =   525
                  Left            =   240
                  Top             =   210
                  Width           =   4680
                  _ExtentX        =   8255
                  _ExtentY        =   926
                  BackColor       =   16761024
                  ForeColor       =   0
                  Caption         =   "【表示中】しばらくお待ち下さい"
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
         End
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
      Height          =   1125
      Left            =   630
      ScaleHeight     =   1065
      ScaleWidth      =   10380
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1710
      Width           =   10440
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   960
         Left            =   6570
         Top             =   60
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   1693
         ForeColor       =   16711680
         Caption         =   "  納 入 日 　  便"
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
         Begin CisText_V60.CisText iH1_NBin 
            Height          =   360
            Left            =   1935
            TabIndex        =   6
            Top             =   435
            Width           =   345
            _ExtentX        =   609
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
            cFbComma        =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
         Begin CisYMD_With_Btn.CisYMDwB iH1_NYmd 
            Height          =   330
            Left            =   165
            TabIndex        =   5
            Top             =   450
            Width           =   1695
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
      Begin Cis3D_v60.CIS3D BLB_HINBN 
         Height          =   960
         Left            =   9045
         Top             =   60
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   1693
         ForeColor       =   16711680
         Caption         =   "手配区分"
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
         Begin VB.ComboBox H1Cmb_TKbn 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   135
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   7
            Top             =   420
            Width           =   1035
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   960
         Left            =   2955
         Top             =   60
         Width           =   2895
         _ExtentX        =   5106
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
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   75
            TabIndex        =   3
            Top             =   435
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
            Left            =   1065
            Top             =   435
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XX"
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
         Left            =   5850
         Top             =   60
         Width           =   720
         _ExtentX        =   1270
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
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   360
            Left            =   165
            TabIndex        =   4
            Top             =   420
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
      Begin Cis3D_v60.CIS3D CIS3D31 
         Height          =   960
         Left            =   45
         Top             =   60
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   1693
         ForeColor       =   16711680
         Caption         =   " 　　発 注 日   発行時間"
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_HYmd 
            Height          =   330
            Left            =   195
            TabIndex        =   0
            Top             =   435
            Width           =   1695
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
         Begin Cis3D_v60.CIS3D Back_HM 
            Height          =   345
            Left            =   1980
            Top             =   435
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   609
            BackColor       =   12640511
            Caption         =   ":"
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
            cAlingnment     =   7
            cBoderStyle     =   1
            Begin CisNfText_V60.CisNfText iH1_HHh 
               Height          =   285
               Left            =   90
               TabIndex        =   1
               Top             =   60
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
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
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "##"
               cILength        =   2
            End
            Begin CisNfText_V60.CisNfText iH1_HMm 
               Height          =   285
               Left            =   480
               TabIndex        =   2
               Top             =   60
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
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
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "##"
               cILength        =   2
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
      Height          =   720
      Left            =   10125
      ScaleHeight     =   660
      ScaleWidth      =   4290
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   825
      Width           =   4350
      Begin CisBtn_60.CisBtn PB_Hin 
         Height          =   585
         Left            =   45
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
         BtnCaption      =   "品番"
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
         PFForeColor     =   65535
         PFShadowColor   =   0
         PFBackColor     =   0
      End
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1725
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
         Left            =   2565
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
         Left            =   3405
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
         Left            =   45
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   885
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
      TabIndex        =   9
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
      Caption         =   "【 かんばん　発注指示  】"
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
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   16711935
         Caption         =   "XXXX"
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
         Left            =   14400
         Top             =   75
         Width           =   915
         _ExtentX        =   1614
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
         Left            =   11850
         Top             =   30
         Width           =   2490
         _ExtentX        =   4392
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
End
Attribute VB_Name = "CXH0030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   かんばん発注指示
'**       フォームID    :   CXH0030
'**       処理概要      :
'**
'**       作  成  日    :   2004/02/02  By CIS
'**       変  更  日    :   2004/04/15  By CIS - 画面中央位置表示の追加
'**       変  更  日    :   2004/05/13  By CIS - 不要部分の削除：iH1_HCYmd（項目自体が無い為）
'**       変  更  日    :   2006/03/10  By CIS - ﾊﾟｽﾜｰﾄﾞ連携対応
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim SyoriKb                 As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM                 As String       ' 処理名称(新規･修正･削除 等)
    Dim CisVsGrid               As New CisVsGrid3
    Dim CisVsGridH              As New CisVsGrid3
    Dim CisComboTehai           As CisCombo
    Dim CisSimplePrint          As New CisSimplePrint
    
    Dim mWhere                  As String
    
    Dim mKaraFlg                As Boolean

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
'   #------------------------------------#
'   # 納入日・便算出方法(システム設定値)
'   #------------------------------------#
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 14
    If Not IRNRead("", 1) Then
        IRN.oyian = "0"
    End If
    If CisFun.Val2(Left(IRN.oyian, 1)) <> 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    納入日･便の算出方法は [発注テーブル] の設定になっています！            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        End
    End If
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 3
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 19
        Set .SelectDispObj = B1lb_SelSu
        .InitGet = False
        .Init
    End With
    Call GridInit
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGridH
        Set .GridObj = VsGrid2
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .InitGet = False
        .Init
    End With
    Call GridInit2
    
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
    '+---------------------+
    '+ ｺﾝﾎﾞﾎﾞｯｸｽ 初期設定  +
    '+---------------------+
    With CisFun
        .INI_Section = "発注処理手配"
        .INI_FileID = gIniExe & "Combo.ini"
        .INI_Default = ""
        .INI_Key = "TYPE"
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_Key = "TYPE"
           .INI_String = "SQL"
           .PutIni
           .INI_Key = "CNT"
           .INI_String = "4"
           .PutIni
           .INI_Key = "DATA1"
           .INI_String = "SELECT 値,値名称 FROM 名称マスタ"
           .PutIni
           .INI_Key = "DATA2"
           .INI_String = "WHERE 区分名称 = '手配区分'   AND 区分タイプ = ''"
           .PutIni
           .INI_Key = "DATA3"
           .INI_String = "And 値 IN('0','1','3')"
           .PutIni
           .INI_Key = "DATA4"
           .INI_String = "ORDER BY 値"
           .PutIni
        End If
    End With
    ' 手配区分
    Set CisComboTehai = New CisCombo
    With CisComboTehai
        Set .Connect = CisDB
        Set .CombObj = H1Cmb_TKbn
        .Control = "発注処理手配"
        .ComboName = ""
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .Ess = False
        .ComboInit
        .ComboSet
    End With
    
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKb = ""
    Call SyoriName(SyoriKb)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)
    
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    FormAct = Not FormAct
    
' 発行時間初期表示
    iH1_HYmd = Format(Date, "YYYYMMDD")
    
    gSL_Select = "Select max(発行時間) 発行時間 from 発注サイクルマスタ "
    gSL_Select = gSL_Select & " Where 発行時間 "
    gSL_Select = gSL_Select & " < Convert(Char(2), Getdate(), 108) + Substring(Convert(Char(5), Getdate(), 108), 4, 2)"
    If HCMRead(gSL_Select, 1) Then
       If HCM.発行時間 <> "0000" Then
          iH1_HHh = Mid(HCM.発行時間, 1, 2)
          iH1_HMm = Mid(HCM.発行時間, 3, 2)
       End If
    End If
' 削除更新済みデータ削除
    With CisDB
        .SQL = "DELETE 発注かんばん予定テーブル "
        .SQL = .SQL & " Where 処理区分 = 2"
        .DBExec
    End With
' 帳票出力テーブル削除
    With CisDB
        .SQL = "Delete 帳票出力テーブル"
        .SQL = .SQL & "    Where 端末番号 = Host_Name()"
        .SQL = .SQL & "            And 帳票区分 IN (12,15)"
        .DBExec
    End With
' 空明細書テーブル削除
    Call MeisyoInit
    
    If Printer.DeviceName = "Mytalk FAX" Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "          通常プリンタがＦＡＸに設定されています              "
            .MB_MSG(4) = "      　通常使うプリンタの変更を手動にて行って下さい。          "
            .MB_Title = ""
            .MB_Button = OK
            .MBOX
        End With
        Call PB_END_Click
    End If
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
           Case vbKeyF1:        If PB_Hin.Visible Then Call PB_Hin_Click    '【品番】
           Case vbKeyF3:        If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
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
'+---------------------+
'+   打切対象品番リスト
'+---------------------+
    Call UchiList
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
   If ProcHB = "H1" Then
      Call HeadBodyClear("H")
      H1Cmb_TKbn.ListIndex = 0
      iH1_HYmd.SetFocus
      GoTo PB_CAN_Ed
   End If
    
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKb)
      iH1_HYmd.SetFocus
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKb <> "DEL" Then
        SyoriKb = "DEL"
        B1lb_SelSu.BackColor = &HC0C0FF
    Else
        SyoriKb = ""
        B1lb_SelSu.BackColor = &HC0FFC0
    End If
    Call SyoriName(SyoriKb)
    iH1_HYmd.SetFocus
End Sub
'+----------------------------+
'+    品番情報表示(F11)       +
'+----------------------------+
Private Sub PB_Hin_Click()
    If B2_Area1.Visible Then
        B2_Area1.Visible = False
    Else
        B2_Area1.Visible = True
        Call GridSet2
    End If
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_HYmd" Then Call iH1_HYmd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_NYmd" Then Call iH1_NYmd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_SSTorcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_UTorcd(0): Exit Sub
    If PB_Look.Tag = "iH1_HHh" Then Call Look_Time: Exit Sub
    If PB_Look.Tag = "iH1_HMm" Then Call Look_Time: Exit Sub
   
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
'+       発行時間検索          +
'+-----------------------------+
Private Sub Look_Time()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    
    CKH0030.Show vbModal
    Unload CKH0030
    Set CKH0030 = Nothing

    If RV_Rtn Then
        If Trim(RV_TorcdK) <> "" Then
           iH1_Torcd = RV_TorcdK
           H1lb_Tornm = RV_TorRName
           iH1_Ukeir = RV_Ukeir
        End If
        If Trim(RV_SYmd) <> "" Then
          iH1_HHh = Mid(RV_SYmd, 1, 2)
          iH1_HMm = Mid(RV_SYmd, 3, 2)
        End If
    End If
End Sub

'+-----------------------------+
'+  iH1_HYmd
'+-----------------------------+
Private Sub iH1_HYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_HYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_HHh
'+-----------------------------+
Private Sub iH1_HHh_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Back_HM.BackColor = iH1_HYmd.cGotColor
    iH1_HMm.cLostColor = iH1_HYmd.cGotColor
End Sub
Private Sub iH1_HHh_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    Back_HM.BackColor = iH1_HYmd.cLostColor
    iH1_HMm.cLostColor = iH1_HYmd.cLostColor
End Sub
'+-----------------------------+
'+  iH1_HMm
'+-----------------------------+
Private Sub iH1_HMm_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Back_HM.BackColor = iH1_HYmd.cGotColor
    iH1_HHh.cLostColor = iH1_HYmd.cGotColor
End Sub
Private Sub iH1_HMm_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    Back_HM.BackColor = iH1_HYmd.cLostColor
    iH1_HHh.cLostColor = iH1_HYmd.cLostColor
    If CisFun.Val2(iH1_HHh) <> 0 Then
       If CisFun.Val2(iH1_HMm) = 0 Then
          iH1_HMm = "00"
       End If
    Else
       iH1_HMm = ""
    End If
End Sub
'+-----------------------------+
'+  iH1_NYmd
'+-----------------------------+
Private Sub iH1_NYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_NYmd_LostFocus()
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
'+  H1Cmb_TKbn
'+-----------------------------+
Private Sub H1Cmb_TKbn_GotFocus()
    H1Cmb_TKbn.BackColor = iH1_HYmd.cGotColor
End Sub
Private Sub H1Cmb_TKbn_LostFocus()
    H1Cmb_TKbn.BackColor = iH1_HYmd.cLostColor
End Sub
''+-----------------------------+
''+  iH1_HCYmd
''+-----------------------------+
'Private Sub iH1_HCYmd_GotFocus()
'    PB_Look.Tag = ActiveControl.Name
'    PB_Look.Visible = True
'End Sub
'Private Sub iH1_HCYmd_LostFocus()
'    If ActiveControl.Name <> "PB_Look" Then
'        PB_Look.Tag = ""
'        PB_Look.Visible = False
'    End If
'End Sub
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
    Static w_Color      As Long
    w_Color = &H80000008 '初期色ｾｯﾄ(黒)
    
    Select Case SyoriKb
        Case "DEL"
                            SyoriNM = "削除"
                            K_Sykbnm.ForeColor = &HFF&              '(赤色)
                            w_Color = &HFF&
        Case Else:          SyoriNM = ""
    End Select
    
'    For Each gObj In Me.Controls
'        If gObj.Name Like "i*" Then gObj.ForeColor = w_Color
'        If gObj.Name Like "*Comb*" Then gObj.ForeColor = w_Color
'
'''       * PF-Key 表示設定 *
''        If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
''           gObj.Name = "PB_DEL" Then
''                gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
''        End If
'    Next gObj
    
    Timer2.Enabled = False
    K_Sykbnm = SyoriNM
    K_Sykbnm.Visible = Not (SyoriNM = "")
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
    If RTrim$(K_Sykbnm) <> "" Then K_Sykbnm = "" Else K_Sykbnm = SyoriNM
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
    
    H1Mode = False:     H1Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    B2Mode = False
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "B1":      B1Mode = True:  B1Color = gPGotSel
        Case "B2":      B2Mode = True
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    
    B2_Area1.Visible = B2Mode
    
'    PB_CAN.Visible = Not H1Mode
    PB_DEL.Visible = H1Mode
    PB_Look.Visible = H1Mode
    PB_Hin.Visible = Not H1Mode
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
    On Error GoTo HeadBodyClear_ER
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then
                .Text = ""
              '  .Enabled = True
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
        CisVsGrid.Clear
    End If
    Exit Sub
HeadBodyClear_ER:
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
       vsGrid1.SetFocus
       GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
       If Not Body1Chk Then GoTo ReturnPress_Ed
    ' << 削除処理 >>
        If SyoriKb = "DEL" Then
            If Not HCDelete Then GoTo ReturnPress_Ed
            Call HeadBodyClear("B")
            ProcHB = "H1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKb)
            GoTo ReturnPress_Ed
        End If
    ' << 発注処理 >>
        If Not HCProc Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKb)
        GoTo ReturnPress_Ed
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
    B1lb_DispSu = ""
    
    mKaraFlg = False
    If Trim(iH1_HYmd) <> "" And CisFun.Val2(iH1_HHh) <> 0 And SyoriKb = "" Then
       mKaraFlg = True
    End If
    ' 空明細 発行
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 13
    Call IRNRead("", 1)
    If Mid(IRN.oyian, 1, 1) <> "1" Then
       mKaraFlg = False
    End If
' 発注かんばん予定テーブル SELECT文 ｾｯﾄ
    If mKaraFlg Then
        ' 発注日チェック
        gInt = Kdou_Check(iH1_HYmd)
        If gInt = 1 Then
        '休日
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(1) = "     発注日が非稼働日です。     "
            CisFun.MB_Button = Error
            CisFun.MBOX
            iH1_HYmd.SetFocus
            Exit Function
        End If
        If gInt = 2 Then
        'カレンダ未登録
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(1) = "     " & RTrim(iH1_HYmd.cYear) & "年" & RTrim(iH1_HYmd.cMonth) & "月のカレンダ未登録！     "
            CisFun.MB_MSG(3) = "     　　(　カレンダマスタ　)　"
            CisFun.MB_Button = Error
            CisFun.MBOX
            iH1_HYmd.SetFocus
            Exit Function
        End If
        
    
        With CisDB
            .SQL = "かんばん発注指示抽出"
            .StoadoCount = 3
            .ParaNo = 0: .ParaIO = Return用
            .ParaNo = 1: .ParaIO = Input用
            .ParaNo = 2: .ParaIO = Input用
            .ParaNo = 3: .ParaIO = OutPut用
        '
            .ParaNo = 1: .ParaValue = iH1_HYmd
            .ParaNo = 2: .ParaValue = Format(iH1_HHh, "00") & Format(iH1_HMm, "00")
        '
            If Not .DBStored Then
               With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    かんばん発注指示抽出エラー            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
               End With
               iH1_HYmd.SetFocus
               Exit Function
            End If
            
            .ParaNo = 0: gCnt2 = .ParaValue
            If gCnt2 <> 0 Then
               With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    かんばん発注指示抽出エラー            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
               End With
               iH1_HYmd.SetFocus
               Exit Function
            End If
        End With
        
        gSL_Select = "Select * From かんばん発注指示ビュー KH "
    Else
        gSL_Select = "Select KH.発行時間,KH.納入日,KH.納入便," & vbCr
        gSL_Select = gSL_Select & " KH.仕入先,KH.受入,KH.手配区分,KH.SK区分," & vbCr
        gSL_Select = gSL_Select & "     KH.発注日,KH.発注便,COUNT(KH.品番) 品番数,SUM(KH.枚数) 枚数," & vbCr
        gSL_Select = gSL_Select & "     CASE WHEN 作成種別 = 'I' Then '一括' Else '指定' END 入力," & vbCr
        gSL_Select = gSL_Select & "     RTRIM(SY.値名称) 手配区分名," & vbCr
        gSL_Select = gSL_Select & "     RTRIM(TR.略称) 仕入先名,作成種別,ISNULL(HT.明細書発行,'1') 明細書発行, " & vbCr
        gSL_Select = gSL_Select & "     CASE When Isnull(HKR.仕入先,'') <> '' Then '済' Else '' End 空発行 " & vbCr
        gSL_Select = gSL_Select & "     From 発注かんばん予定テーブル KH" & vbCr
        gSL_Select = gSL_Select & "    Left Outer Join" & vbCr
        gSL_Select = gSL_Select & "         ( SELECT 発注日,発注便,納入日,納入便,発行時間," & vbCr
        gSL_Select = gSL_Select & "                仕入先 , 受入" & vbCr
        gSL_Select = gSL_Select & "                 From 空明細書テーブル" & vbCr
        gSL_Select = gSL_Select & "                 Where 発行区分 = ''" & vbCr
        gSL_Select = gSL_Select & "                 Group by 発注日,発注便,納入日,納入便,発行時間," & vbCr
        gSL_Select = gSL_Select & "                        仕入先,受入 ) HKR " & vbCr
        gSL_Select = gSL_Select & "          ON HKR.発注日 = KH.発注日" & vbCr
        gSL_Select = gSL_Select & "         And HKR.発注便 = KH.発注便" & vbCr
        gSL_Select = gSL_Select & "         And HKR.発行時間 = KH.発行時間" & vbCr
        gSL_Select = gSL_Select & "         And HKR.納入日 = KH.納入日" & vbCr
        gSL_Select = gSL_Select & "         And HKR.納入便 = KH.納入便" & vbCr
        gSL_Select = gSL_Select & "         And HKR.仕入先 = KH.仕入先" & vbCr
        gSL_Select = gSL_Select & "         And HKR.受入   = KH.受入" & vbCr
        gSL_Select = gSL_Select & "     Left Outer Join 品番取引先マスタ HT" & vbCr
        gSL_Select = gSL_Select & "         ON  HT.品番 = KH.品番" & vbCr
        gSL_Select = gSL_Select & "         And HT.取引先 = KH.仕入先" & vbCr
        gSL_Select = gSL_Select & "         And HT.受入 = KH.受入" & vbCr
        gSL_Select = gSL_Select & "         And HT.品目 = 1" & vbCr
        gSL_Select = gSL_Select & "     Left Outer Join 名称マスタ SY" & vbCr
        gSL_Select = gSL_Select & "         ON  SY.区分名称 = '手配区分'" & vbCr
        gSL_Select = gSL_Select & "         And SY.区分タイプ = ''" & vbCr
        gSL_Select = gSL_Select & "         And SY.値 = KH.手配区分" & vbCr
        gSL_Select = gSL_Select & "     Left Outer Join 取引先マスタ TR" & vbCr
        gSL_Select = gSL_Select & "         ON  TR.取引先区分 = 1" & vbCr
        gSL_Select = gSL_Select & "         And TR.取引先CD = KH.仕入先" & vbCr
    End If
    ' 抽出条件
    mWhere = ""
    With CisFun
        ' 発注日･発行時間
        .Where_CreateKB = NewDocument
        If Not mKaraFlg Then
           .Where_Create Numeric, "処理区分", "0", , Create
        End If
        .Where_Create AlphaMe, "KH.発注日", iH1_HYmd
        If CisFun.Val2(iH1_HHh) <> 0 Or _
           CisFun.Val2(iH1_HMm) <> 0 Then
           .Where_Create AlphaMe, "KH.発行時間", Format(iH1_HHh, "00") & Format(iH1_HMm, "00")
        End If
        ' 納入日･便
        .Where_Create AlphaMe, "KH.納入日", iH1_NYmd
        If CisFun.Val2(iH1_NBin) <> 0 Then
           .Where_Create AlphaMe, "KH.納入便", Format(iH1_NBin, "00")
        End If
        ' 仕入先
        .Where_Create AlphaMe, "KH.仕入先", iH1_Torcd
        ' 受入
        .Where_Create AlphaMe, "KH.受入", iH1_Ukeir
        ' 手配区分
        If Trim(H1Cmb_TKbn) <> "" Then
            Call CisComboTehai.CodeGet
            .Where_Create AlphaMe, "KH.手配区分", CisComboTehai.Code
        End If
        ' 発注日･便
'        .Where_Create AlphaMe, "発注日", iH1_HCYmd
'        If CisFun.Val2(iH1_HCBin) <> 0 Then
'           .Where_Create AlphaMe, "発注便", Format(iH1_HCBin, "00")
'        End If
        gSL_Select = gSL_Select & .Where_Phrase
    End With
    
    If mKaraFlg Then
        gSL_Select = gSL_Select & " Order By 発注日,発注便,発行時間,納入日,納入便," & vbCr
        gSL_Select = gSL_Select & "            仕入先,受入,手配区分,作成種別" & vbCr
    Else
        gSL_Select = gSL_Select & "     Group by KH.発注日,KH.発注便,KH.発行時間,KH.納入日,KH.納入便," & vbCr
        gSL_Select = gSL_Select & "            KH.仕入先,KH.受入,KH.手配区分,KH.SK区分," & vbCr
        gSL_Select = gSL_Select & "            KH.作成種別," & vbCr
        gSL_Select = gSL_Select & "            SY.値名称 , TR.略称,HT.明細書発行,HKR.仕入先" & vbCr
        
        gSL_Select = gSL_Select & " Order By KH.発注日,KH.発注便,KH.発行時間,KH.納入日,KH.納入便," & vbCr
        gSL_Select = gSL_Select & "            KH.仕入先,KH.受入,KH.手配区分,KH.SK区分," & vbCr
        gSL_Select = gSL_Select & "            KH.作成種別" & vbCr
    End If
    If Not HKYRead(gSL_Select) Then
       Call HKYClose
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    該当データがありません            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
       End With
       iH1_HYmd.SetFocus
       Exit Function
    End If
    
    Message.Visible = True
    DoEvents
    
    Call GridSet
    
    Message.Visible = False
    
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    
    Body1Chk = False
    
    If CisVsGrid.SelectCount = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    処理対象行を選択して下さい            "
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
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+      削 除 処 理      +
'+----------------------+
Private Function HCDelete()
    Dim YDelFlg     As Boolean
    HCDelete = False
'  ( 確認メッセージ )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     削除処理を行います      "
        .MB_Title = "削除処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    On Error GoTo HCDelete_Err
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
    YDelFlg = False
    With vsGrid1
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(3))) = "" Then Exit For

            If Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(CisVsGrid.SelectCol))) = "X" Then
                If CisFun.Val2(.TextMatrix(gLong, CisVsGrid.FixedGet(9))) = 0 Then
                Else
                    ' 発注かんばん予定テーブル 削除（処理区分=2 にて更新)
                    CisDB.SQL = "UPDATE 発注かんばん予定テーブル "
                    CisDB.SQL = CisDB.SQL & " Set 処理区分 = 2,"
                    CisDB.SQL = CisDB.SQL & " 処理端末 = Host_Name() "
                    If UpdateWhere(gLong, "") Then
                        CisDB.SQL = CisDB.SQL & mWhere
                        If Not CisDB.DBExec Then GoTo HCDelete_Err
                        YDelFlg = True
                    End If
                End If
            End If
        Next gLong
        If YDelFlg Then
           If Not KanabnDel Then GoTo HCDelete_Err
        End If
    End With
    
    Call CisDB.DBTran(TransCommit)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    HCDelete = True
    Exit Function
HCDelete_Err:
    Call CisDB.DBTran(TransRollback)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 破棄
    CisFun.ErrorBox
End Function

'//////////////////////////////////////
'//
'//     発注かんばん予定削除 更新
'//
'//////////////////////////////////////
Private Function KanabnDel() As Boolean
    KanabnDel = False
    
    On Error GoTo KanabnDel_ER
    gStr = ""
    With CisDB
        .SQL = "発注かんばん予定削除処理"
        .StoadoCount = 2
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = OutPut用
        .ParaNo = 2: .ParaIO = OutPut用
        
        If Not .DBStored Then GoTo KanabnDel_ER
        
        ' ストアド処理エラー
        .ParaNo = 0: gCnt1 = CisFun.Val2(.ParaValue)
        If gCnt1 <> 0 Then GoTo KanabnDel_ER
        
    End With

    KanabnDel = True
    On Error GoTo 0
    Exit Function
KanabnDel_ER:
    Select Case gLong
        Case 1, 2, 3: gStr = "かんばんチェックテーブル 更新エラー"
        Case 4: gStr = "発注打切テーブル 更新エラー"
        Case 5: gStr = "帳票出力テーブル 作成エラー"
        Case 6: gStr = "帳票出力テーブル 作成エラー"
        Case Else
             gStr = "その他ｴﾗｰ:" & gStr
    End Select
    gStr = StrConv(Format(gLong, "00"), vbWide) & ":" & gStr
    Call CisFun.ErrorBox(gStr)

End Function
'+----------------------+
'+    発注処理 確認     +
'+----------------------+
Private Function HCProc() As Boolean
    HCProc = False
    Dim AddFlg  As Boolean

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     発注処理を行います      "
        .MB_Title = "発注処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo HCProc_Er

    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' 帳票出力テーブル削除
    With CisDB
        .SQL = "Delete 帳票出力テーブル"
        .SQL = .SQL & "    Where 端末番号 = Host_Name()"
        .SQL = .SQL & "            And 帳票区分 = 15"
        .DBExec
    End With
    
    With vsGrid1
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(3))) = "" Then Exit For

            If Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(CisVsGrid.SelectCol))) = "X" Then
                If CisFun.Val2(.TextMatrix(gLong, CisVsGrid.FixedGet(9))) = 0 Then
                    ' 空明細テーブル作成
                    If Not KaraMeiPut(gLong, "") Then GoTo HCProc_Er
                Else
                    ' 空明細テーブル作成(発注あり）
                    If Not KaraMeiPut(gLong, "H") Then GoTo HCProc_Er
                    ' 発注かんばん予定テーブル (処理区分=1 にて更新)
                    CisDB.SQL = "UPDATE 発注かんばん予定テーブル Set "
                    CisDB.SQL = CisDB.SQL & " 処理区分 = 1,"
                    CisDB.SQL = CisDB.SQL & " 処理端末 = Host_Name() "
                    If UpdateWhere(gLong, "") Then
                        CisDB.SQL = CisDB.SQL & mWhere
                        CisDB.DBExec
                    End If
                End If
            End If
        Next gLong
    End With
    
    Call CisDB.DBTran(TransCommit)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    ' 発注処理
    '------------------------ 2006/03/10 insert
    Call CisDB.PassLink
    '------------------------ 2006/03/10 insert End
    With CisFun
        .ShellApl = gIniExe & "PXH0110.exe"
        .ShellAplPara = "Y"
        .ShellTask NoFreeze
    End With

    HCProc = True
    Exit Function
HCProc_Er:
    Call CisDB.DBTran(TransRollback)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 破棄
    CisFun.ErrorBox
End Function
'+--------------------------------------------+
'+      空明細テーブル作成
'+--------------------------------------------+
Private Function KaraMeiPut(GId As Long, OutType As String) As Boolean
    KaraMeiPut = False
    
    On Error GoTo KaraMeiPut_er
    
    Call ItemsClearKMT
    
    If OutType = "" Then
        With CisDB
        '   <><><> 明細書管理番号の獲得 <><><>
            .SQL = "番号獲得"
            .StoadoCount = 3
            .ParaNo = 0: .ParaIO = Return用
            .ParaNo = 1: .ParaIO = Input用
            .ParaNo = 2: .ParaIO = OutPut用
            .ParaNo = 3: .ParaIO = Return用
        '
            .ParaNo = 1: .ParaValue = 5
        '
            If Not .DBStored Then
                CisFun.MB_Lines = 4
                CisFun.MB_MSG(2) = "    明細書管理番号の獲得に失敗しました            "
                CisFun.MB_Title = "採番処理"
                CisFun.MB_Button = Error
                CisFun.MBOX
                Exit Function
            Else
                .ParaNo = 2: KMT.明細書NO = .ParaValue
            End If
        End With
    Else
        gSL_Select = "Select * From 空明細書テーブル"
        gSL_Select = gSL_Select & " Where 明細書NO = 0"
        gSL_Select = gSL_Select & " And   発行区分 = '" & RTrim(OutType) & "'"
        gSL_Select = gSL_Select & " And   仕入先 = '" & Trim(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(4))) & "'"
        gSL_Select = gSL_Select & " And   受入 = '" & Trim(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(6))) & "'"
        gSL_Select = gSL_Select & " And   納入日 = '" & Trim(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(13))) & "'"
        gSL_Select = gSL_Select & " And   納入便 = '" & Trim(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(14))) & "'"
        gSL_Select = gSL_Select & " And   発注日 = '" & Trim(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(15))) & "'"
        gSL_Select = gSL_Select & " And   発注便 = '" & Trim(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(16))) & "'"
        gSL_Select = gSL_Select & " And   発行時間 = '" & Format(iH1_HHh, "00") & Format(iH1_HMm, "00") & "'"
        If KMTRead(gSL_Select, 1) Then
           KaraMeiPut = True
           Exit Function
        End If
        Call ItemsClearKMT
    End If
    
    With KMT
        .発行区分 = OutType
        .仕入先 = Trim(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(4)))
        .受入 = Trim(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(6)))
        .納入日 = Trim(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(13)))
        .納入便 = Trim(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(14)))
        .発注日 = Trim(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(15)))
        .発注便 = Trim(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(16)))
        .発行時間 = Format(iH1_HHh, "00") & Format(iH1_HMm, "00")
        .仕入先名 = Trim(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(5)))
        .作成者 = gTanto
    End With
    If Not KMTInsert Then GoTo KaraMeiPut_er
    
    If OutType <> "" Then
       KaraMeiPut = True
       Exit Function
    End If
    ' 帳票出力テーブル出力
    Call ItemsClearCHT
    CHT.端末番号 = "' + substring(host_name(),1,20) + '"
    CHT.帳票区分 = 15
    CHT.管理文字 = Trim(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(18)))
    CHT.管理NO = KMT.明細書NO
    If Not CHTInsert Then GoTo KaraMeiPut_er
    
    KaraMeiPut = True
    Exit Function
KaraMeiPut_er:

End Function
'+--------------------------------------------+
'+ 発注かんばん予定テーブル 更新用 Where作成
'+--------------------------------------------+
Private Function UpdateWhere(wRow As Long, CrtType As String) As Boolean
    UpdateWhere = False
    
    mWhere = ""
    
    With vsGrid1
        ' 明細書発行
        gStr = RTrim(.TextMatrix(wRow, CisVsGrid.FixedGet(18)))
        If Trim(CrtType) = "" Then
            
            mWhere = mWhere & " From 品番取引先マスタ HT"
            mWhere = mWhere & "   Where HT.品番 = 発注かんばん予定テーブル.品番"
            mWhere = mWhere & "     And HT.取引先 = 発注かんばん予定テーブル.仕入先"
            mWhere = mWhere & "     And HT.受入    = 発注かんばん予定テーブル.受入"
            mWhere = mWhere & "     And HT.品目   = 1"
            mWhere = mWhere & "     And HT.明細書発行 = '" & gStr & "'"
        Else
            mWhere = mWhere & " Where HT.明細書発行 = '" & gStr & "'"
        End If
'        ' 発注日
        gStr = RTrim(.TextMatrix(wRow, CisVsGrid.FixedGet(15)))
        mWhere = mWhere & " And 発注日 = '" & gStr & "'"
        ' 発行時間
        gStr = Replace(.TextMatrix(wRow, CisVsGrid.FixedGet(2)), ":", "")
        mWhere = mWhere & " And 発行時間 = '" & gStr & "'"
        ' 納入日
        gStr = RTrim(.TextMatrix(wRow, CisVsGrid.FixedGet(13)))
        mWhere = mWhere & " And 納入日 = '" & gStr & "'"
        ' 納入便
        gStr = RTrim(.TextMatrix(wRow, CisVsGrid.FixedGet(14)))
        mWhere = mWhere & " And 納入便 = '" & gStr & "'"
        ' 仕入先
        gStr = RTrim(.TextMatrix(wRow, CisVsGrid.FixedGet(4)))
        mWhere = mWhere & " And 仕入先 = '" & gStr & "'"
        ' 受入
        gStr = RTrim(.TextMatrix(wRow, CisVsGrid.FixedGet(6)))
        mWhere = mWhere & " And 発注かんばん予定テーブル.受入 = '" & gStr & "'"
        ' 手配区分
        gStr = RTrim(.TextMatrix(wRow, CisVsGrid.FixedGet(12)))
        mWhere = mWhere & " And 発注かんばん予定テーブル.手配区分 = '" & gStr & "'"
        ' 発注便
        gStr = RTrim(.TextMatrix(wRow, CisVsGrid.FixedGet(16)))
        mWhere = mWhere & " And 発注便 = '" & gStr & "'"
        ' 作成種別
        gStr = RTrim(.TextMatrix(wRow, CisVsGrid.FixedGet(17)))
        mWhere = mWhere & " And 作成種別 = '" & gStr & "'"
        ' 処理区分
        mWhere = mWhere & " And 処理区分 <> 2"
    End With

    If mWhere <> "" Then
'       mWhere = " Where " & mWhere
       UpdateWhere = True
    End If
End Function
'+----------------------------------+
'+  打切対象品番削除リスト 作表
'+----------------------------------+
Private Sub UchiList()

    gSL_Select = "SELECT ISNULL(TR.略称,'') 仕入先名,"
    gSL_Select = gSL_Select & " CASE WHEN Isnull(HN.表示品番,'') = '' THEN KY.品番 Else HN.表示品番 End 表示品番,"
    gSL_Select = gSL_Select & "  KY.品番,KY.納入日,KY.納入便,KY.収容数,SUM(KY.枚数) 枚数,SUM(発注数) 発注数, "
    gSL_Select = gSL_Select & "  KY.仕入先"
    gSL_Select = gSL_Select & "  From 帳票出力テーブル CH"
    gSL_Select = gSL_Select & "  Left Outer Join 発注かんばん予定テーブル KY"
    gSL_Select = gSL_Select & "      ON  KY.管理番号 = CH.管理NO"
    gSL_Select = gSL_Select & "  Left Outer Join 品番マスタ HN"
    gSL_Select = gSL_Select & "      ON  HN.品番 = KY.品番"
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & "      ON  TR.取引先CD = KY.仕入先"
    gSL_Select = gSL_Select & "      And TR.取引先区分 = 1"
    gSL_Select = gSL_Select & " Where CH.端末番号 = Host_Name()"
    gSL_Select = gSL_Select & "   And CH.帳票区分 = 12"
    gSL_Select = gSL_Select & "  Group by KY.品番,KY.納入日,KY.納入便,KY.収容数,"
    gSL_Select = gSL_Select & "           KY.仕入先,TR.略称,HN.表示品番"
    gSL_Select = gSL_Select & "  Order By KY.品番,KY.仕入先,KY.納入日,KY.納入便"
    If Not HKYRead(gSL_Select) Then
       Call HKYClose
       Exit Sub
    End If


    With CisSimplePrint
        .PrintTitle = "※ 打切対象品番削除リスト ※"
        .Preservation = App.Title
        .PrintSize = A4
        .PrintOrientation = Portrait
        .PrintHeadLine = あり
        .PrintBodyGyo = 10
        
        .pHeadItemSet 0, 左, "№"
        .pHeadItemSet 1, 中, "品      番"
        .pHeadItemSet 2, 中, "仕   入   先"
        .pHeadItemSet 3, 中, "納 入 日"
        .pHeadItemSet 4, 中, "便"
        .pHeadItemSet 5, 右, "収容数"
        .pHeadItemSet 6, 右, "枚数"
        .pHeadItemSet 7, 右, "発注数"
    
        .pBodyItemSet 0, 2, 右, "##"
        .pBodyItemSet 1, 25
        .pBodyItemSet 2, 25
        .pBodyItemSet 3, 11
        .pBodyItemSet 4, 2
        .pBodyItemSet 5, 7, 右, "#,###"
        .pBodyItemSet 6, 5, 右, "#,###"
        .pBodyItemSet 7, 7, 右, "#,###"
        
        .BodyInit
        
        .MaxCount = CisDB.RecordCount
        gLong = 0
        Do Until Not HKY_RDSTS
            gLong = gLong + 1
            .pBodyValueSet 0, Format(gLong, "0")
            .pBodyValueSet 1, HKY.表示品番
            .pBodyValueSet 2, HKY.仕入先 & ":" & HKY.仕入先名
            CisFun.DateE_BefVal = HKY.納入日: CisFun.DateEdit
            .pBodyValueSet 3, CisFun.DateE_AfVal
            .pBodyValueSet 4, HKY.納入便
            .pBodyValueSet 5, HKY.収容数
            .pBodyValueSet 6, HKY.枚数
            .pBodyValueSet 7, HKY.発注数
            .BodySet
        
            Call HKYReadNext
        Loop
        
        .PrintStatus = gDemoPrt + 1
        
        .PrintStart
    
    End With
End Sub
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理　関　連
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
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
        Call .FixedSet(中中, 中中, あり, 14, 1, "発 注 日 - 便")
        Call .FixedSet(中中, 中中, あり, 6, 2, "時間")
        Call .FixedSet(中中, 中中, あり, 14, 3, "納 入 日 - 便")
        Call .FixedSet(中中, 左中, あり, 7, 4, "仕入先", 数字)
        Call .FixedSet(中中, 左中, あり, 26, 5, "仕入先名")
        Call .FixedSet(中中, 中中, あり, 4, 6, "受入")
        Call .FixedSet(中中, 中中, あり, 5, 7, "手配")
        Call .FixedSet(中中, 中中, あり, 0, 8, "出力")
        Call .FixedSet(中中, 右中, あり, 5, 9, "枚数")
        Call .FixedSet(中中, 中中, あり, 5, 10, "入力")
        Call .FixedSet(中中, 中中, あり, 6, 11, "空明細")
        Call .FixedSet(中中, 左中, あり, 0, 12, "手配区分")
        Call .FixedSet(中中, 左中, あり, 0, 13, "納入日")
        Call .FixedSet(中中, 左中, あり, 0, 14, "納入便")
        Call .FixedSet(中中, 左中, あり, 0, 15, "発注日")
        Call .FixedSet(中中, 左中, あり, 0, 16, "発注便")
        Call .FixedSet(中中, 左中, あり, 0, 17, "作成種別")
        Call .FixedSet(中中, 中中, あり, 11, 18, "明細書発行")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 19, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,9,10,11,8")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long
    
    CisVsGrid.Clear
' 抽出処理 初期ｾｯﾄ
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
    B1lb_DispSu = Format(Ttl_Kensu, "#,###")
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid1
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(ID)
        Next ID
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not HKY_RDSTS
        With vsGrid1
            ID = ID + 1
            '№
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            ' 発注日･便
            If Trim(HKY.発注日) <> "" Then
               With CisFun
                    .DateE_BefVal = HKY.発注日
                    .DateEdit
                    gStr = .DateE_AfVal
               End With
               If Trim(HKY.発注便) <> "" Then
                  gStr = gStr & "-" & Format(HKY.発注便, "00")
               Else
                  gStr = gStr & "   "
               End If
               .TextMatrix(ID, GridCol(1)) = gStr
               .TextMatrix(ID, GridCol(15)) = HKY.発注日
               .TextMatrix(ID, GridCol(16)) = HKY.発注便
            End If
            ' 発行時間
            .TextMatrix(ID, GridCol(2)) = Format(Mid(HKY.発行時間, 1, 2), "00") & ":" & _
                                          Format(Mid(HKY.発行時間, 3, 2), "00")
            ' 納入日･便
            If Trim(HKY.納入日) <> "" Then
               With CisFun
                    .DateE_BefVal = HKY.納入日
                    .DateEdit
                    gStr = .DateE_AfVal
               End With
               If Trim(HKY.納入便) <> "" Then
                  gStr = gStr & "-" & Format(HKY.納入便, "00")
               Else
                  gStr = gStr & "   "
               End If
               .TextMatrix(ID, GridCol(3)) = gStr
            
               .TextMatrix(ID, GridCol(13)) = HKY.納入日
               .TextMatrix(ID, GridCol(14)) = HKY.納入便
            End If
            ' 仕入先
            .TextMatrix(ID, GridCol(4)) = HKY.仕入先
            .TextMatrix(ID, GridCol(5)) = HKY.仕入先名
            ' 受入
            .TextMatrix(ID, GridCol(6)) = HKY.受入
            ' 手配先
            .TextMatrix(ID, GridCol(7)) = HKY.手配区分名
            .TextMatrix(ID, GridCol(12)) = HKY.手配区分
            ' 品番数
'            .TextMatrix(ID, GridCol(8)) = Format(HKY.品番数, "###")
            ' 出力
            .TextMatrix(ID, GridCol(9)) = Format(HKY.枚数, "###")
            '出力
            Select Case HKY.明細書発行
                Case "1": gStr = ""
                Case "3": gStr = "FAX"
                Case "4": gStr = "Mail"
                Case Else
                    gStr = "？？"
            End Select
            .TextMatrix(ID, GridCol(8)) = gStr
            ' 入力
            .TextMatrix(ID, GridCol(10)) = HKY.入力
            ' 空明細
            If HKY.品番数 = 0 And HKY.枚数 = 0 Then
               .TextMatrix(ID, GridCol(11)) = "○"
            End If
            If HKY.空発行 = "済" Then
               .TextMatrix(ID, GridCol(11)) = "済"
            End If
            '
            .TextMatrix(ID, GridCol(17)) = HKY.作成種別
'            If Trim(HKY.発行日) <> "" Then
'               With CisFun
'                    .DateE_BefVal = HKY.発行日
'                    .DateEdit
'                    gStr = .DateE_AfVal
'               End With
'               .TextMatrix(ID, GridCol(18)) = gStr
'            End If
            .TextMatrix(ID, GridCol(18)) = HKY.明細書発行
       End With
        Call HKYReadNext
    Loop
    Call HKYClose
    With vsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'******************************************'
'*****        行選択（クリック）       *****
'******************************************'
Private Sub VSGrid1_Click()
    If SyoriKb = "" Then
       Call CisVsGrid.vsColor
    Else
       Call CisVsGrid.vsColor(, &HC0C0FF)
    End If
End Sub
'******************************************'
'*****        行選択（スペースｷｰ)      *****
'******************************************'
Private Sub VSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    Call VSGrid1_Click
End Sub

Private Sub vsGrid1_RowColChange()
    Call GridSet2
End Sub
'******************************************'
'*****       移動選択項目選択          *****
'******************************************'
Private Sub VSGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsGrid1
        .Tag = ""
    
        If Button <> vbLeftButton Then Exit Sub
        If .MouseRow <> 0 Or .MouseCol = 0 Then Exit Sub
    
        .Tag = Str(.MouseCol)
        .Drag vbBeginDrag
    End With
End Sub
'******************************************'
'*****        表 示 項 目 移 動        *****
'******************************************'
Private Sub VSGrid1_DragDrop(Source As Control, X As Single, Y As Single)
    With vsGrid1
        If .MouseCol = 0 Then Exit Sub
        If .Tag = "" Then Exit Sub
    
        .Redraw = False
    
        .ColPosition(CisFun.Val2(.Tag)) = .MouseCol
        .Col = 1: .ColSel = .Cols - 1
    
        .Redraw = True
    
        CisVsGrid.VSSort
    End With
End Sub
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit2()
    With CisVsGridH
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 0, 0, "No")
        Call .FixedSet(中中, 左中, あり, 25, 1, "品     番")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
        'Call .FixedSet(中中, 中中, あり, 6, 2, "背番")
        Call .FixedSet(中中, 中中, あり, 8, 2, "背番号")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
        Call .FixedSet(中中, 中中, あり, 8, 3, "サイクル")
        Call .FixedSet(中中, 右中, あり, 6, 4, "収容数")
        Call .FixedSet(中中, 右中, あり, 4, 5, "枚数")
        Call .FixedSet(中中, 右中, あり, 8, 6, "発注数")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
'        Call .FixedSet(中中, 左中, なし, 0,3,"Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6")
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット(品番情報)*****
'******************************************'
Private Sub GridSet2()
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long

    If Not B2_Area1.Visible Then Exit Sub

' 抽出処理 初期ｾｯﾄ
    CisVsGridH.Clear
    B2lb_Nymd = "": B2lb_Torcd = ""
    
    If Trim(vsGrid1.TextMatrix(vsGrid1.Row, CisVsGrid.FixedGet(3))) = "" Then
       Exit Sub
    End If
    
    With vsGrid1
        B2lb_Nymd = .TextMatrix(.Row, CisVsGrid.FixedGet(3))
        B2lb_Torcd = .TextMatrix(.Row, CisVsGrid.FixedGet(4)) & "-" & _
                     .TextMatrix(.Row, CisVsGrid.FixedGet(6)) & ":" & _
                     .TextMatrix(.Row, CisVsGrid.FixedGet(5))
    End With
    
    Call UpdateWhere(vsGrid1.Row, "X")
    mWhere = Replace(mWhere, "発注かんばん予定テーブル", "HY")
    gSL_Select = "Select 表示品番,HY.背番号,HY.収容数,HY.サイクル,SUM(HY.枚数) 枚数,SUM(HY.発注数) 発注数 "
    gSL_Select = gSL_Select & " From 発注かんばん予定テーブル HY "
    gSL_Select = gSL_Select & " Left Outer Join 品番マスタ HN "
    gSL_Select = gSL_Select & "    ON HN.品番 = HY.品番"
    gSL_Select = gSL_Select & " Left Outer Join 品番取引先マスタ HT"
    gSL_Select = gSL_Select & "   ON  HT.品番 = HY.品番"
    gSL_Select = gSL_Select & "   And HT.取引先 = HY.仕入先"
    gSL_Select = gSL_Select & "   And HT.受入    = HY.受入"
    gSL_Select = gSL_Select & "   And HT.品目   = 1"
    gSL_Select = gSL_Select & mWhere
    gSL_Select = gSL_Select & " Group By 表示品番,HY.背番号,HY.サイクル,HY.収容数"
    gSL_Select = gSL_Select & " Order By 表示品番"
    If Not HKYRead(gSL_Select) Then
       Call HKYClose
       Exit Sub
    End If
    
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGrid2
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGridH.FixedGet(ID)
        Next ID
        If Ttl_Kensu + 1 < CisVsGridH.DispMax Then
            .Rows = CisVsGridH.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not HKY_RDSTS
        With VsGrid2
            ID = ID + 1
            '№
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            '項目1
            .TextMatrix(ID, GridCol(1)) = HKY.表示品番
            .TextMatrix(ID, GridCol(2)) = HKY.背番号
            If Trim(HKY.サイクル) <> "" Then
               .TextMatrix(ID, GridCol(3)) = Mid(HKY.サイクル, 1, 1) & "-" & _
                                             Mid(HKY.サイクル, 2, 2) & "-" & _
                                             Mid(HKY.サイクル, 4, 2)
            End If
            .TextMatrix(ID, GridCol(4)) = Format(HKY.収容数, "#,###")
            .TextMatrix(ID, GridCol(5)) = Format(HKY.枚数, "#,###")
            .TextMatrix(ID, GridCol(6)) = Format(HKY.発注数, "#,###")
       End With
        Call HKYReadNext
    Loop
    Call HKYClose
    With VsGrid2
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'+---------------------------------------+
'+      空明細書テーブル削除
'+---------------------------------------+
Private Sub MeisyoInit()
    Dim wKDel           As Long
    Dim tDay            As String
' 空明細書削除日を求める
    gSL_Select = " SELECT * FROM 名称マスタ "
    gSL_Select = gSL_Select & " WHERE 区分名称 = '空明細書削除管理'"
    gSL_Select = gSL_Select & "   AND 区分タイプ = '' "
    gSL_Select = gSL_Select & "   AND 値 = '1' "
    If Not SYMRead(gSL_Select, 1) Then
        Call ItemsClearSYMCreate
        With SYMCreate
            .区分名称 = "空明細書削除管理"
            .区分桁数 = 1
            .値 = "1"
            .値名称 = "空明細書削除日数"
            .数字1 = 10
            .作成者 = gTanto
        End With
        Call SYMInitCreate(True)
        SYM.数字1 = 10
    End If
    wKDel = SYM.数字1
    If wKDel = 0 Then Exit Sub
    tDay = ""
    With CisDB
        .SQL = "日付加算"
        .StoadoCount = 4
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = Input用
        .ParaNo = 3: .ParaIO = OutPut用
        .ParaNo = 4: .ParaIO = OutPut用

        .ParaNo = 1: .ParaValue = ""
        If wKDel <> 0 Then wKDel = wKDel * -1
        .ParaNo = 2: .ParaValue = wKDel

        If .DBStored Then
           .ParaNo = 4: gLong = CisFun.Val2(.ParaValue)
           If gLong = 0 Then
              .ParaNo = 3: tDay = Trim(.ParaValue)
           End If
        End If
        If Trim(tDay) <> "" Then
           .SQL = "Delete 空明細書テーブル Where 納入日 < '" & tDay & "'"
           .DBExec
        End If
    End With

End Sub

