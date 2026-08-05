VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CXM1030 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "かんばんポケット№メンテナンス"
   ClientHeight    =   10980
   ClientLeft      =   1980
   ClientTop       =   1950
   ClientWidth     =   15315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   10980
   ScaleWidth      =   15315
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D CIS3D3 
      Height          =   9090
      Left            =   13935
      Top             =   1335
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   16034
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   240
         Left            =   90
         Top             =   75
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   423
         Caption         =   "ポケット№"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
         cAlingnment     =   6
         cBoderStyle     =   2
      End
      Begin vsFlexLib.vsFlexArray vsGrid2 
         Height          =   8640
         Left            =   90
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   330
         Width           =   975
         _Version        =   196608
         _ExtentX        =   1720
         _ExtentY        =   15240
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
         FormatString    =   "   |   "
         Rows            =   30
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   16777215
         BackColorSel    =   16761024
         BackColorBkg    =   14737632
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   0
         RowHeightMin    =   230
         GridLines       =   2
         SelectionMode   =   1
         BackColorAlternate=   16777215
      End
   End
   Begin Cis3D_v60.CIS3D CIS3D4 
      Height          =   750
      Left            =   10425
      Top             =   570
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   1323
      BackColor       =   0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
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
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   705
         Left            =   30
         Top             =   30
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1244
         Caption         =   "表示件数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cPositionX      =   50
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_Kensu 
            Height          =   285
            Left            =   60
            Top             =   360
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            BackColor       =   16777152
            Caption         =   "##,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
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
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   7365
      Left            =   525
      ScaleHeight     =   7305
      ScaleWidth      =   13275
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3045
      Width           =   13335
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   7200
         Left            =   60
         Top             =   60
         Width           =   13170
         _ExtentX        =   23230
         _ExtentY        =   12700
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
            Height          =   7035
            Left            =   90
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   60
            Width           =   12945
            _Version        =   196608
            _ExtentX        =   22834
            _ExtentY        =   12409
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
            FormatString    =   "№  |仕入先              |受|品目|品番                          |背番号|収容数|サイクル|PNO|設定|X仕入先|X品番|X品目|Select|"
            Rows            =   21
            Cols            =   16
            BackColor       =   12648384
            BackColorSel    =   16761024
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   0
            RowHeightMin    =   330
            GridLines       =   2
            SelectionMode   =   1
            Begin CisText_V60.CisText iB1_SetSu 
               Height          =   375
               Left            =   11430
               TabIndex        =   9
               Top             =   330
               Visible         =   0   'False
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   661
               cFormat         =   "#0"
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
               cFbComma        =   0
               cGFormat        =   "#0"
               cILength        =   2
            End
            Begin CisText_V60.CisText iB1_BinSu 
               Height          =   375
               Left            =   12000
               TabIndex        =   10
               Top             =   330
               Visible         =   0   'False
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   661
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
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "###"
               cILength        =   3
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
      Height          =   1290
      Left            =   525
      ScaleHeight     =   1230
      ScaleWidth      =   13290
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1320
      Width           =   13350
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   0
         Left            =   2250
         Top             =   60
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1984
         ForeColor       =   16711680
         Caption         =   "品目"
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
         Begin VB.OptionButton H1Op_Hinmoku 
            BackColor       =   &H00C0FFC0&
            Caption         =   "発注"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   180
            TabIndex        =   8
            Top             =   750
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton H1Op_Hinmoku 
            BackColor       =   &H00C0C0C0&
            Caption         =   "受注"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   7
            Top             =   390
            Width           =   795
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   3
         Left            =   7155
         Top             =   60
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1984
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
         cPositionX      =   30
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   360
            Left            =   150
            TabIndex        =   1
            Top             =   510
            Width           =   375
            _ExtentX        =   661
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
      Begin Cis3D_v60.CIS3D Back_Torcd 
         Height          =   1125
         Left            =   3495
         Top             =   60
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   1984
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
         cPositionX      =   30
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   90
            TabIndex        =   0
            Top             =   510
            Width           =   1005
            _ExtentX        =   1773
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
            Left            =   1080
            Top             =   510
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   635
            BackColor       =   16777152
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
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   4
         Left            =   8385
         Top             =   60
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   1984
         ForeColor       =   16711680
         Caption         =   "品　　番"
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
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   360
            Left            =   90
            TabIndex        =   3
            Top             =   510
            Width           =   3195
            _ExtentX        =   5636
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
            IMEMode         =   2
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   5
         Left            =   7815
         Top             =   60
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   1984
         ForeColor       =   16711680
         Caption         =   "PNO"
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
         Begin CisText_V60.CisText iH1_PNO 
            Height          =   360
            Left            =   90
            TabIndex        =   2
            Top             =   510
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   635
            cFormat         =   "#0"
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
            cGFormat        =   "#0"
            cILength        =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   1984
         ForeColor       =   16711680
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   60
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   795
            Left            =   60
            Top             =   150
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "表　示"
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
         Begin VB.OptionButton H1Op_Disp 
            BackColor       =   &H00C0C0C0&
            Caption         =   "品番"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   390
            TabIndex        =   6
            Top             =   780
            Width           =   1755
         End
         Begin VB.OptionButton H1Op_Disp 
            BackColor       =   &H00C0FFC0&
            Caption         =   "仕入先"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   390
            TabIndex        =   4
            Top             =   90
            Value           =   -1  'True
            Width           =   1755
         End
         Begin VB.OptionButton H1Op_Disp 
            BackColor       =   &H00C0C0C0&
            Caption         =   "仕入先・受入"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   390
            TabIndex        =   5
            Top             =   420
            Width           =   1755
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   6
         Left            =   11715
         Top             =   60
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1984
         ForeColor       =   16711680
         Caption         =   "表示順"
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
         Begin VB.OptionButton H1Op_SortKb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "取引先"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   390
            Width           =   1320
         End
         Begin VB.OptionButton H1Op_SortKb 
            BackColor       =   &H00C0FFC0&
            Caption         =   "ﾎﾟｹｯﾄ№"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   750
            Value           =   -1  'True
            Width           =   1320
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
      Left            =   11580
      ScaleHeight     =   660
      ScaleWidth      =   3435
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   555
      Width           =   3495
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   900
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
         Left            =   1740
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
         Left            =   2580
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
   End
   Begin Cis3D_v60.CIS3D UC_3D11 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【かんばんポケット№メンテナンス 】"
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
         Left            =   14385
         Top             =   60
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
         Height          =   300
         Left            =   11835
         Top             =   30
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   529
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXXXXXXX2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   0
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
      Left            =   705
      Top             =   -120
   End
   Begin VB.Timer Timer1 
      Left            =   270
      Top             =   -150
   End
   Begin VB.TextBox Dummy 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Left            =   -15
      TabIndex        =   12
      Text            =   "Dummy"
      Top             =   -105
      Width           =   345
   End
End
Attribute VB_Name = "CXM1030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  かんばんポケット№メンテナンス
'**       フォームID    :  CXM1030
'**       処理概要      :
'**
'**       作  成  日    :  2005/05/16
'**       変  更  日    :   2005/09/30  表示順の追加・設定済ﾎﾟｹｯﾄ№の表示
'**       変  更  日    :   2006/02/03 By CIS ポケット最大数を越えたポケットの色設定
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl

    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim CisVsGrid1              As New CisVsGrid3
    Dim CisVsGrid2              As New CisVsGrid3
'   品目の背景色
    Dim mColor_On               As Long
    Dim mColor_Off              As Long
'   グリッドの背景色
    Dim mColorOrg               As Long
    Dim mColorChg               As Long
'   グリッド2の背景色
    Dim mGrid2_On               As Long
    Dim mGrid2_Off              As Long
    Dim mGrid2_OnOff            As Long
    Dim mGrid2_ErrP             As Long
    
    Const mColorFix             As Long = &HFFC0C0
'   行
    Dim NowRow                  As Long
    Dim NewRow                  As Long

    Dim MaxPocket               As Long
    
    Dim mMaxPoketColor          As Long             ' 2006/02/03 Insert
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》            +
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

'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid1
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .EditCheckCol = 1
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 15
        .InitGet = False
        .Init
    End With
    Call GridInit
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid2
        Set .GridObj = vsGrid2
        .NoSpaceCol = 1
        .NumberSet = No
        .EditCheckCol = 1
        .FixedHeight = 250
        .PatanMax = 1
        .SelectCol = 31
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
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》           +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'   背景色の設定（グリッド）
    mColorOrg = vsGrid1.BackColor
    mColorChg = &HC0E0FF
'   背景色の設定（グリッド2）
    mGrid2_On = &HFF8080
    mGrid2_Off = vsGrid2.BackColor
    mGrid2_OnOff = &HC0E0FF
    mGrid2_ErrP = &HFFC0FF
'   背景色の設定（表示・品目）
    mColor_On = H1Op_Hinmoku(1).BackColor
    mColor_Off = H1Op_Hinmoku(0).BackColor
    H1Op_Disp(0).Value = True
    H1Op_Hinmoku(1).Value = True
    H1Op_SortKb(1).Value = True
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》               +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》            +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》    +
'+----------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Key_Acc = False

    Select Case Key_Code

        Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click      '【検索】
        Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click        '【取消】
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click        '【終了】
        Case vbKeyReturn
                            If PB_Ent.Visible Then                          '【入力】
                                Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                Call PB_ENT_Click
                                Dummy.Enabled = False
                            End If
        Case Else:          Exit Function
    End Select

    Key_Acc = True
End Function
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    Call CisDB.DBDISConnect
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
       GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_Ukeir: Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_PNO" Then Call Look_PNo: Exit Sub
End Sub
'+----------------------------+
'+      取引先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 2
    
    If H1Op_Hinmoku(0).Value Then
       RV_TorKb = 3
    Else
       RV_TorKb = 8
    End If
        
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
    RV_TorKb = 0
    RV_Torcd = iH1_Torcd
    RV_Kojyo = ""

    If H1Op_Hinmoku(0).Value Then
       RV_TorKb = 3
    Else
       RV_TorKb = 8
    End If

    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
        iH1_Ukeir = RV_Ukeir
    End If
End Sub
'+----------------------------+
'+      品番検索              +
'+----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
    RVI_Torcd = RTrim(iH1_Torcd)
    RVI_Ukeir = RTrim(iH1_Ukeir)
    RVI_Hinmoku = 9
    If H1Op_Hinmoku(0).Value Then
       RVI_Hinmoku = 0
    Else
       RVI_Hinmoku = 1
    End If

    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iH1_Hinbn = RV_Hinbn
        iH1_Torcd = RV_TorcdK
        iH1_Ukeir = RV_Ukeir
        H1lb_Tornm = ""
        gInt = 0
        If H1Op_Hinmoku(0).Value Then
           gInt = 3
        Else
           gInt = 9
        End If
        If TorNmGet(iH1_Torcd, gInt) Then
           H1lb_Tornm = TRM_RName
        End If
    End If
End Sub
'+----------------------------+
'+      PNO検索            +
'+----------------------------+
Private Sub Look_PNo()
    RV_Left = 0
    RV_Top = 0
    
    RV_Torcd = ""
    If H1Op_Hinmoku(0).Value Then
       RV_TorKb = 0
    Else
       RV_TorKb = 1
    End If
        
    CXM1035.Show vbModal
    Unload CXM1035
    Set CXM1035 = Nothing

    If RV_Rtn Then
       iH1_PNO = CisFun.RSetFld(RV_Code, 2, "#0")
    End If
End Sub
'+-----------------------------+
'+  H1Op_Disp
'+-----------------------------+
Private Sub H1Op_Disp_Click(Index As Integer)
    For Each gObj In H1Op_Disp
        If Index = gObj.Index Then
           gObj.BackColor = mColor_On
        Else
           gObj.BackColor = mColor_Off
        End If
    Next gObj
    
    If Index = 0 Then
        iH1_Ukeir = "": iH1_Ukeir.Enabled = False
        iH1_Hinbn = "": iH1_Hinbn.Enabled = False
    End If
    If Index = 1 Then
        iH1_Ukeir.Enabled = True
        iH1_Hinbn = "": iH1_Hinbn.Enabled = False
    End If
    If Index = 2 Then
        iH1_Ukeir.Enabled = True
        iH1_Hinbn.Enabled = True
    End If

    vsGrid1.Redraw = False
    Call GridInit
    vsGrid1.Redraw = True
    
End Sub
'+-----------------------------+
'+  H1Op_Hinmoku
'+-----------------------------+
Private Sub H1Op_Hinmoku_Click(Index As Integer)
    If Index = 0 Then
        H1Op_Hinmoku(1).BackColor = mColor_Off
        H1Op_Hinmoku(0).BackColor = mColor_On
        H1Op_Disp(0).Caption = "納入先"
        H1Op_Disp(1).Caption = "納入先・受入"
        Back_Torcd.Caption = "納入先"
        H1Op_SortKb(0).Caption = "納入先"
    Else
        H1Op_Hinmoku(0).BackColor = mColor_Off
        H1Op_Hinmoku(1).BackColor = mColor_On
        H1Op_Disp(0).Caption = "仕入先"
        H1Op_Disp(1).Caption = "仕入先・受入"
        Back_Torcd.Caption = "仕入先"
        H1Op_SortKb(0).Caption = "仕入先"
    End If
    Call GridInit2
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
    gInt = 0
    If H1Op_Hinmoku(0).Value <> H1Op_Hinmoku(1).Value Then
       If H1Op_Hinmoku(0).Value = 1 Then
          gInt = 3
       Else
          gInt = 9
       End If
    End If
    If TorNmGet(iH1_Torcd, gInt) Then
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
'+  iH1_PNO
'+-----------------------------+
Private Sub iH1_PNO_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_PNO_LostFocus()
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
'+  H1Op_SortKb
'+-----------------------------+
Private Sub H1Op_SortKb_Click(Index As Integer)
    If Index = 0 Then
        H1Op_SortKb(1).BackColor = mColor_Off
        H1Op_SortKb(0).BackColor = mColor_On
    Else
        H1Op_SortKb(0).BackColor = mColor_Off
        H1Op_SortKb(1).BackColor = mColor_On
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

    H1Mode = False: H1Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1": H1Mode = True: H1Color = gPGotSel
       Case "B1": B1Mode = True: B1Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode: H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode: B1_Area1.BackColor = B1Color

    PB_Look.Visible = False
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode

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
    If HeadBody = "B" Then
        Call CisVsGrid1.Clear
        Call GridInit2
        iB1_SetSu.Visible = False
        iB1_BinSu.Visible = False
    End If
End Sub
'*********************************
'*          入力処理(Enter)      *
'*********************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        Call DispChange(ProcHB)
        Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
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
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False

    gSL_Select = "Select HT.PNO,"
    
    If H1Op_Disp(0).Value Then
        gSL_Select = gSL_Select & "     HT.取引先,ISNULL(TR.略称,'') 取引先名,Count(*) 件数"
    End If
    If H1Op_Disp(1).Value Then
        gSL_Select = gSL_Select & "     HT.取引先,HT.受入,ISNULL(TR.略称,'') 取引先名,Count(*) 件数"
    End If
    If H1Op_Disp(2).Value Then
        gSL_Select = gSL_Select & " HT.品目,"
        gSL_Select = gSL_Select & " Case When isnull(HT.変更日, '') = '' then 収容数1"
        gSL_Select = gSL_Select & "          When HT.変更日 > Convert(Char(8), GetDate(), 112)"
        gSL_Select = gSL_Select & "                   then HT.収容数1"
        gSL_Select = gSL_Select & "                   Else: HT.収容数2"
        gSL_Select = gSL_Select & "        end 収容数1,"
        gSL_Select = gSL_Select & "     HT.サイクル,"
        gSL_Select = gSL_Select & "     HN.表示品番,HT.品番,HT.背番号,"
        gSL_Select = gSL_Select & "     HT.取引先,HT.受入,ISNULL(TR.略称,'') 取引先名"
    End If
    
    gSL_Select = gSL_Select & "     From 品番取引先マスタ HT"
    gSL_Select = gSL_Select & "     Left Outer Join 品番マスタ HN"
    gSL_Select = gSL_Select & "         ON HN.品番 = HT.品番"
    gSL_Select = gSL_Select & "     Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & "         ON  TR.取引先区分 = HT.品目"
    gSL_Select = gSL_Select & "         And TR.取引先CD = HT.取引先"
    
    
    With CisFun
        .Where_CreateKB = NewDocument
        .Where_Create AlphaMe, "HT.取引先", iH1_Torcd
        .Where_Create AlphaMe, "HT.受入", iH1_Ukeir
        .Where_Create AlphaMe, "HT.品番", iH1_Hinbn
        .Where_Create Numeric, "HT.SK区分", 0, 等しい, Create
        If Trim(iH1_PNO) <> "" Then
           .Where_Create Numeric, "HT.PNO", iH1_PNO, 等しい, Create
        End If
        If H1Op_Hinmoku(0).Value Then
          .Where_Create Numeric, "HT.品目", 0, 等しい, Create
        Else
          .Where_Create Numeric, "HT.品目", 1
        End If
        gSL_Select = gSL_Select & .Where_Phrase
    End With
    If H1Op_Disp(0).Value Then
       gSL_Select = gSL_Select & " Group by HT.PNO,HT.取引先,ISNULL(TR.略称,'') "
        If H1Op_SortKb(0).Value Then
           gSL_Select = gSL_Select & " Order By HT.取引先,HT.PNO"
        Else
           gSL_Select = gSL_Select & " Order By HT.PNO,HT.取引先 "
        End If
    End If
    If H1Op_Disp(1).Value Then
       gSL_Select = gSL_Select & " Group by HT.PNO,HT.取引先,HT.受入,ISNULL(TR.略称,'') "
        If H1Op_SortKb(0).Value Then
           gSL_Select = gSL_Select & " Order By HT.取引先,HT.受入,HT.PNO "
        Else
           gSL_Select = gSL_Select & " Order By HT.PNO,HT.取引先,HT.受入 "
        End If
    End If
    If H1Op_Disp(2).Value Then
        If H1Op_SortKb(0).Value Then
           gSL_Select = gSL_Select & " Order By HT.取引先,HT.受入,HT.品目,HT.品番,HT.PNO "
        Else
           gSL_Select = gSL_Select & " Order By HT.PNO,HT.取引先,HT.受入,HT.品目,HT.品番 "
        End If
    End If
    
    
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
    Body1Chk = False

    Call CisVsGrid1.EditEnable
    Call ChgColorGrid1(NowRow)      ' 内容変更ﾁｪｯｸ

    Dim ID          As Long
    With vsGrid1
        For ID = 1 To .Rows - 1
            If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(0))) = "" Then Exit For
        '   設定枚数
            If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(9))) > MaxPocket Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    " & ID & "行目 ： 設定枚数が最大ポケット№を超えています            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                .Row = ID
                Call CisVsGrid1.EditEnable
                Call CisVsGrid1.EditVisible
                Exit Function
            End If
        Next ID
    End With

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
    'Call GridInit2
    
    If GridSet1 Then
        DBInput = True
    End If

DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False

    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "    更新処理を行います。            "
        .MB_MSG(4) = "     よろしいですか？            "
        .MB_Title = "更新処理"
        .MB_Button = OK_CAN
        If Not .MBOX Then
           vsGrid1.Col = 0: vsGrid1.CellBackColor = mColorFix
           Exit Function
        End If
    End With
    On Error GoTo DBPut_Err

    Call CisDB.DBTran(TransBegin)   ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

' ( 更新処理 )
    Dim ID          As Long
    With vsGrid1
        For ID = 1 To .Rows - 1
            If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(0))) = "" Then Exit For
            If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(9))) <> "" Then
                Call DBUpdate_HTM(ID)
            End If
        Next ID
    End With

    CisDB.DBTran (TransCommit)      ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+------------------------------------+
'+      品番取引先マスタ更新
'+------------------------------------+
Public Function DBUpdate_HTM(SetRow As Long, Optional DBNo As Integer) As Boolean
    DBUpdate_HTM = False

    Dim wTorcd      As String
    Dim wUkeir      As String
    Dim wHinbn      As String
    Dim wPNo        As Long
    Dim wHNo        As Long

    Dim wHinmoku    As Byte
    
    With vsGrid1
        wTorcd = Trim(.TextMatrix(SetRow, CisVsGrid1.FixedGet(10)))
        wUkeir = Trim(.TextMatrix(SetRow, CisVsGrid1.FixedGet(2)))
        wHinbn = Trim(.TextMatrix(SetRow, CisVsGrid1.FixedGet(11)))
        wPNo = CisFun.Val2(.TextMatrix(SetRow, CisVsGrid1.FixedGet(8)))
        wHNo = CisFun.Val2(.TextMatrix(SetRow, CisVsGrid1.FixedGet(9)))
        wHinmoku = 0
        If H1Op_Hinmoku(1).Value Then
           wHinmoku = 1
        End If
    
        gSL_Select = "Update 品番取引先マスタ Set "
        gSL_Select = gSL_Select & "PNO = " & wHNo & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(gTanto) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        ' 取引先
        If H1Op_Disp(0).Value Then
           gSL_Select = gSL_Select & " Where 取引先 = '" & wTorcd & "'"
           gSL_Select = gSL_Select & "   And PNO = " & wPNo
           gSL_Select = gSL_Select & "   And 品目 = " & wHinmoku
           gSL_Select = gSL_Select & "   And SK区分 = 0"
        End If
        ' 取引先･受入
        If H1Op_Disp(1).Value Then
           gSL_Select = gSL_Select & " Where 取引先 = '" & wTorcd & "'"
           gSL_Select = gSL_Select & "   And 受入 = '" & wUkeir & "'"
           gSL_Select = gSL_Select & "   And PNO = " & wPNo
           gSL_Select = gSL_Select & "   And 品目 = " & wHinmoku
           gSL_Select = gSL_Select & "   And SK区分 = 0"
        End If
        ' 品番
        If H1Op_Disp(2).Value Then
           gSL_Select = gSL_Select & " Where 品番 = '" & wHinbn & "'"
           gSL_Select = gSL_Select & "   And 取引先 = '" & wTorcd & "'"
           gSL_Select = gSL_Select & "   And 受入 = '" & wUkeir & "'"
           gSL_Select = gSL_Select & "   And SK区分 = 0"
        End If
        
    End With

    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    DBUpdate_HTM = True
End Function



'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理　関　連
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 4, 0, "№")
        Call .FixedSet(中中, 左中, あり, 30, 1, "取 引 先")
        Call .FixedSet(中中, 左中, あり, 2, 2, "受")
        Call .FixedSet(中中, 中中, あり, 5, 3, "品目")
        Call .FixedSet(中中, 左中, あり, 30, 4, "品番")
        Call .FixedSet(中中, 左中, あり, 6, 5, "背番号")
        Call .FixedSet(中中, 右中, あり, 10, 6, "収容数")
        If H1Op_Disp(2).Value Then
           Call .FixedSet(中中, 中中, あり, 8, 7, "サイクル")
        Else
           Call .FixedSet(中中, 右中, あり, 8, 7, "件数")
        End If
        Call .FixedSet(中中, 右中, あり, 4, 8, "PNO")
        Call .FixedSet(中中, 右中, あり, 4, 9, "設定")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, あり, 0, 10, "X仕入先")
        Call .FixedSet(中中, 左中, あり, 0, 11, "X品番")
        Call .FixedSet(中中, 左中, あり, 0, 12, "X品目")
        Call .FixedSet(中中, 右中, あり, 0, 13, "Sort")
        Call .FixedSet(中中, 右中, あり, 0, 14, "PNO退避")
        Call .FixedSet(中中, 左中, あり, 0, 15, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,4,5,6,7,8,9")
'   編集内容ｾｯﾄ
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO,非表示ｶﾗﾑ,編集ﾌｫｰﾏｯﾄ,日付編集)
        Call .EditSet(iB1_SetSu, 8, 0, "#0")
        .InitDisp
    End With
End Sub
'***********************************************'
'*****      グリッド内容セット
'***********************************************'
Private Function GridSet1() As Boolean
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

    GridSet1 = False

    If HTMRead(gSL_Select, , 1) Then

'       抽出処理 初期ｾｯﾄ
        B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
        Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
'       ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
        With vsGrid1
           .Redraw = False
           .Rows = Ttl_Kensu + 1
        End With
        ID = 0
        Do Until Not HTM_RDSTS
            ID = ID + 1

            Call InfoToGrid1(ID)

            Call HTMReadNext(1)
        Loop
        With vsGrid1
             .Row = 1
             
            If H1Op_SortKb(1).Value Then
                .Col = CisVsGrid1.FixedGet(13)
                .ColSel = CisVsGrid1.FixedGet(13)
                .Sort = 1
                Call CisVsGrid1.ReNumber
                .Row = 1
            End If
            .Col = 0: .CellBackColor = mColorFix
            .Col = 1: .ColSel = .Cols - 1
            If Ttl_Kensu + 1 < CisVsGrid1.DispMax Then
                .Rows = CisVsGrid1.DispMax
            End If
            .Redraw = True
        End With
        NowRow = vsGrid1.Row
        

        GridSet1 = True
    End If
    Call HTMClose(1)
End Function
'+----------------------------------------------+
'+  品番取引先マスタよりグリッドに転送
'+----------------------------------------------+
Private Sub InfoToGrid1(SetRow As Long)
    With vsGrid1
        'No
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(0)) = SetRow
        '仕入先
        If RTrim(HTM.取引先名) = "" Then
           .TextMatrix(SetRow, CisVsGrid1.FixedGet(1)) = RTrim(HTM.取引先)
        Else
           .TextMatrix(SetRow, CisVsGrid1.FixedGet(1)) = RTrim(HTM.取引先) & ":" & RTrim(HTM.取引先名)
        End If
        '受入
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(2)) = RTrim(HTM.受入)
        ' 品目
        If HTM.品目 = 0 Then
           .TextMatrix(SetRow, CisVsGrid1.FixedGet(3)) = "受注"
        Else
           .TextMatrix(SetRow, CisVsGrid1.FixedGet(3)) = "発注"
        End If
        '品番
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(4)) = RTrim(HTM.表示品番)
        '背番号
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(5)) = RTrim(HTM.背番号)
        '収容数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(6)) = Format(HTM.収容数1, "#,###")
        'サイクル
        If H1Op_Disp(2).Value Then
           If Trim(HTM.サイクル) <> "" Then
              .TextMatrix(SetRow, CisVsGrid1.FixedGet(7)) = Mid(HTM.サイクル, 1, 1) & "-" & _
                                                            Mid(HTM.サイクル, 2, 2) & "-" & _
                                                            Mid(HTM.サイクル, 4, 2)
           End If
        Else
           .TextMatrix(SetRow, CisVsGrid1.FixedGet(7)) = Format(HTM.件数, "#,###")
        End If
        'PNO
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(8)) = Format(HTM.PNO, "##")
        
        
        '仕入先(取引先)
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(10)) = HTM.取引先
        '品番
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(11)) = HTM.品番
        '品目
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(12)) = HTM.品目
        'Grid2 Color
        If HTM.PNO <> 0 Then
            Call vsGrid2_Color(0, HTM.PNO, mGrid2_On)
        End If
        'PNO
'        .TextMatrix(SetRow, CisVsGrid1.FixedGet(14)) = Format(HTM.PNO, "##")
        'Sort-Key
        If HTM.PNO = 0 Then
            HTM.PNO = 99
        End If
        If H1Op_SortKb(1).Value Then
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(13)) = Format(HTM.PNO, "00") & RTrim(HTM.取引先) & Space(7 - CisFun.Len2(HTM.取引先)) & RTrim(HTM.受入) & Space(2 - CisFun.Len2(HTM.受入)) & RTrim(HTM.品番) & Space(25 - CisFun.Len2(HTM.品番))
        End If
    End With
End Sub
'+-------------------------------+
'+      VSGrid1_Click
'+-------------------------------+
Private Sub VSGrid1_Click()
    If ProcHB = "B1" Then

        If Trim(vsGrid1.TextMatrix(vsGrid1.Row, CisVsGrid1.FixedGet(0))) = "" Then
            vsGrid1.Row = NowRow
        End If
        Call ChgColorGrid1(NowRow)      ' 内容変更ﾁｪｯｸ
        Call CisVsGrid1.EditEnable
        Call CisVsGrid1.EditVisible
        If Trim(vsGrid1.TextMatrix(vsGrid1.Row, CisVsGrid1.FixedGet(0))) <> "" Then
            NowRow = vsGrid1.Row
        End If
        vsGrid1.Col = 0: vsGrid1.CellBackColor = mColorFix
    End If
End Sub
'+---------------------------------------------------------+
'+          グリッド編集エリア
'+---------------------------------------------------------+
'+-------------------------------+
'+      iB1_SetSu
'+-------------------------------+
Private Sub iB1_SetSu_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call BD_InKeyCheck(ActiveControl, KeyCode, Shift)
End Sub
'+-------------------------------+
'+      iB1_BinSu
'+-------------------------------+
Private Sub iB1_BinSu_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call BD_InKeyCheck(ActiveControl, KeyCode, Shift)
End Sub
'+--------------------------------------+
'+      ボディ入力項目キーチェック
'+--------------------------------------+
Private Function BD_InKeyCheck(TObj As Object, KeyCode As Integer, Shift As Integer)

'< PageUp・PageDown >
    If KeyCode = vbKeyPageDown Or _
       KeyCode = vbKeyPageUp Then
        
        If vsGrid1.Rows <= CisVsGrid1.DispMax Then Exit Function
        
        Call ChgColorGrid1(NowRow)      ' 内容変更ﾁｪｯｸ
        
        If KeyCode = vbKeyPageDown Then
            
            '< PageDown >
            If vsGrid1.TopRow + 25 <= vsGrid1.Rows - 1 Then
                vsGrid1.TopRow = vsGrid1.TopRow + 25
                
                If vsGrid1.TopRow + 24 <= vsGrid1.Rows - 1 Then
                    vsGrid1.Row = vsGrid1.TopRow + 24
                Else
                    vsGrid1.Row = vsGrid1.Rows - 1
                End If
            Else
                vsGrid1.TopRow = vsGrid1.Rows - 1 - 24
            End If
        Else
            
            '< PageUp >
            If vsGrid1.TopRow - 25 > 0 Then
                vsGrid1.TopRow = vsGrid1.TopRow - 25
            Else
                vsGrid1.TopRow = 1
            End If
            vsGrid1.Row = vsGrid1.TopRow
        End If
        
        vsGrid1.SetFocus
        Call CisVsGrid1.EditEnable
        Call CisVsGrid1.EditVisible
        NowRow = vsGrid1.Row
        vsGrid1.Col = 0: vsGrid1.CellBackColor = mColorFix
        TObj.SetFocus
        Exit Function
    End If
    
'< ↑・↓ >
    If KeyCode = vbKeyUp Or _
       KeyCode = vbKeyDown Then
        Call ChgColorGrid1(NowRow)      ' 内容変更ﾁｪｯｸ
        vsGrid1.SetFocus
        Call CisVsGrid1.KeyEdit(TObj, KeyCode, Shift)
        vsGrid1.Col = 0: vsGrid1.CellBackColor = mColorFix
        NowRow = vsGrid1.Row
    End If

End Function
'+-------------------------------+
'+      背景色の再設定
'+-------------------------------+
Private Sub ChgColorGrid1(SetRow As Long)
    With vsGrid1
        .Redraw = False
        NewRow = .Row
        .Row = SetRow
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(9)) = iB1_SetSu             '空白と０は違う意味
        If CisFun.Val2(iB1_SetSu) = CisFun.Val2(.TextMatrix(SetRow, CisVsGrid1.FixedGet(8))) Or _
           CisFun.Val2(iB1_SetSu) = 0 Then
            For gInt = 1 To .Cols - 1
                .Col = gInt
                .CellBackColor = mColorOrg
            Next gInt
            If RTrim(iB1_SetSu) <> RTrim(.TextMatrix(SetRow, CisVsGrid1.FixedGet(14))) Then
                Call vsGrid2_ColorAll
            End If
        Else
            For gInt = 1 To .Cols - 1
                .Col = gInt
                .CellBackColor = mColorChg
            Next gInt
            If RTrim(iB1_SetSu) <> RTrim(.TextMatrix(SetRow, CisVsGrid1.FixedGet(14))) Then
                Call vsGrid2_ColorAll
            End If
        End If
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(14)) = iB1_SetSu
        .Col = 0: .CellBackColor = &H8000000F
        .Row = NewRow
        .Redraw = True
    End With
End Sub
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit2()
    CisVsGrid2.Clear

'ソーター最大ポケット数
    mMaxPoketColor = 0                          ' 2006/02/03 Insert
    If MeisyoGet("ポケット最大数", "1") Then
        MaxPocket = SYM.数字1
        mMaxPoketColor = SYM.数字2              ' 2006/02/03 Insert
    Else
        MaxPocket = 60
    End If
'№表示
    gCnt2 = 0
    With vsGrid2
        .Redraw = False
        For gInt = 0 To 1
            For gCnt1 = 0 To 29
                gCnt2 = gCnt2 + 1
                If gCnt2 <= MaxPocket Then
                    vsGrid2.Row = gCnt1
                    vsGrid2.Col = gInt
                    vsGrid2.CellAlignment = flexAlignRightCenter
                    .TextMatrix(gCnt1, gInt) = Format(gCnt2, "##")
'------------------------------------------------------------------------ 2006/02/03 Insert
                Else
                    If mMaxPoketColor <> 0 Then
                       .Row = gCnt1
                       .Col = gInt
                       .CellBackColor = mMaxPoketColor
                    End If
                End If
'------------------------------------------------------------------------ 2006/02/03 Insert  End
            Next gCnt1
        Next gInt
    
        .Redraw = True
    End With
'エラーポケット№
    Call vsGrid2_ErrPocket
        
End Sub
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub vsGrid2_Color(vs2Kb As Byte, vs2PNO As Byte, vs2Color As Long)
    
    If vs2PNO = 0 Then Exit Sub
    If vs2PNO > MaxPocket Then Exit Sub
    If vs2PNO Mod 30 <> 0 Then
        vsGrid2.Col = vs2PNO \ 30
        vsGrid2.Row = (vs2PNO Mod 30) - 1
    Else
        vsGrid2.Col = (vs2PNO \ 30) - 1
        vsGrid2.Row = vsGrid2.Rows - 1
    End If
    
    gLong = vsGrid2.CellBackColor
    vsGrid2.CellBackColor = vs2Color
    If vs2Kb = 0 And mGrid2_On = vs2Color Then
        If mGrid2_On <> gLong Then
            vsGrid2.TextMatrix(vsGrid2.Row, vsGrid2.Col) = "*" & vsGrid2.TextMatrix(vsGrid2.Row, vsGrid2.Col)
        End If
    End If
        
End Sub
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub vsGrid2_ColorAll()
    
    With vsGrid2
        .Redraw = False
    '色初期化
        For gCnt1 = 1 To 60
            Call vsGrid2_Color(1, CByte(gCnt1), mGrid2_Off)
        Next
    End With
    'エラーポケット№
    Call vsGrid2_ErrPocket
    With vsGrid1
    '設定済ポケット№
        For gInt = 1 To .Rows - 1
            If CisFun.Val2(.TextMatrix(gInt, CisVsGrid1.FixedGet(8))) <> 0 And RTrim(.TextMatrix(gInt, CisVsGrid1.FixedGet(9))) = "" Then
                Call vsGrid2_Color(1, CisFun.Val2(.TextMatrix(gInt, CisVsGrid1.FixedGet(8))), mGrid2_On)
            End If
        Next gInt
    '今回設定ポケット№
        For gInt = 1 To .Rows - 1
            If RTrim(.TextMatrix(gInt, CisVsGrid1.FixedGet(9))) <> "" Then
                Call vsGrid2_Color(1, CisFun.Val2(.TextMatrix(gInt, CisVsGrid1.FixedGet(9))), mGrid2_OnOff)
            End If
        Next gInt
    End With
    vsGrid2.Redraw = True
        
End Sub
'+----------------------------------------------------------+
'+   エラーポケットの色設定
'+----------------------------------------------------------+
Private Sub vsGrid2_ErrPocket()
    
    If Not H1Op_Hinmoku(0).Value Then
        gSL_Select = "SELECT 値 FROM 名称マスタ"
        gSL_Select = gSL_Select & " WHERE 区分名称 = 'かんばんエラー設定発注'"
        gSL_Select = gSL_Select & "   AND 区分タイプ = ''"
        If SYMRead(gSL_Select, 0) Then
            Do Until Not SYM_RDSTS
                Call vsGrid2_Color(1, CisFun.Val2(SYM.値), &HFFC0FF)
                Call SYMReadNext
            Loop
        End If
        Call SYMClose
    Else
        With vsGrid2
            For gCnt1 = 0 To 3
                .Col = 0: .Row = gCnt1
                .CellBackColor = &HFFC0FF
            Next
        End With
    End If
        
End Sub

