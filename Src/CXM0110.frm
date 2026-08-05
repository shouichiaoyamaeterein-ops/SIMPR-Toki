VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXM0110 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "材料・品番関連付け"
   ClientHeight    =   11025
   ClientLeft      =   -255
   ClientTop       =   945
   ClientWidth     =   15345
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
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
   ScaleHeight     =   11025
   ScaleWidth      =   15345
   Begin VB.PictureBox Back_Daiyo 
      Height          =   2670
      Left            =   2400
      ScaleHeight     =   2610
      ScaleWidth      =   3600
      TabIndex        =   20
      Top             =   3840
      Width           =   3660
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   2550
         Left            =   45
         Top             =   0
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   4498
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
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   405
            Left            =   30
            Top             =   45
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   714
            BackColor       =   16761087
            ForeColor       =   0
            Caption         =   "代用(ｶｯﾄ)材料寸法"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin vsFlexLib.vsFlexArray VsGrid2 
            Height          =   2040
            Left            =   30
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   450
            Width           =   3465
            _Version        =   196608
            _ExtentX        =   6112
            _ExtentY        =   3598
            _StockProps     =   228
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "  |板厚   |幅     |長さ   "
            Rows            =   6
            Cols            =   5
            BackColor       =   12632319
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   0
            RowHeightMin    =   330
            SelectionMode   =   1
            Begin CisText_V60.CisText iB1_DItatu 
               Height          =   360
               Left            =   315
               TabIndex        =   22
               Top             =   315
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   635
               cFormat         =   "###0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###0.00"
               MaxLength       =   7
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   2
               cFdAutoFormat   =   1
               cGFormat        =   "###0.00"
               cILength        =   4
            End
            Begin CisText_V60.CisText iB1_DWidth 
               Height          =   360
               Left            =   1260
               TabIndex        =   23
               Top             =   315
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   635
               cFormat         =   "###0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###0.00"
               MaxLength       =   7
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   2
               cFdAutoFormat   =   1
               cGFormat        =   "###0.00"
               cILength        =   4
            End
            Begin CisText_V60.CisText iB1_DLong 
               Height          =   360
               Left            =   2205
               TabIndex        =   24
               Top             =   315
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   635
               cFormat         =   "###0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###0.00"
               MaxLength       =   7
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   2
               cFdAutoFormat   =   1
               cGFormat        =   "###0.00"
               cILength        =   4
            End
         End
      End
   End
   Begin VB.Timer Timer4 
      Left            =   210
      Top             =   810
   End
   Begin VB.PictureBox B2_Area1 
      Height          =   7245
      Left            =   450
      ScaleHeight     =   7185
      ScaleWidth      =   14385
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3480
      Width           =   14445
      Begin vsFlexLib.vsFlexArray VsGrid1 
         Height          =   7020
         Left            =   90
         TabIndex        =   13
         Top             =   90
         Width           =   14205
         _Version        =   196608
         _ExtentX        =   25056
         _ExtentY        =   12382
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   "No |品番|背番号|材質|寸法|ピッチ|取数|使用|屑重量|展|ＬＴ|有効期限|材料管理番号|編集前品番|有効期限開始|有効期限終了|単位||"
         Rows            =   21
         Cols            =   23
         BackColor       =   12648447
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   330
         SelectionMode   =   1
         AllowBigSelection=   0   'False
         BackColorAlternate=   12648447
         AllowSelection  =   0   'False
      End
   End
   Begin VB.PictureBox B1_Area1 
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
      Height          =   1215
      Left            =   450
      ScaleHeight     =   1155
      ScaleWidth      =   14385
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2160
      Width           =   14445
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   1005
         Left            =   12120
         Top             =   90
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   1773
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
         cPositionY      =   30
         Begin CisYMD_With_Btn.CisYMDwB iB1_EDate 
            Height          =   330
            Left            =   570
            TabIndex        =   12
            Top             =   615
            Width           =   1545
            _ExtentX        =   2249
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
            Object.Width           =   1545
            cSize           =   1
            cChkResult      =   0   'False
            cCalenderPosition=   1
            cType           =   1
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
         Begin CisYMD_With_Btn.CisYMDwB iB1_SDate 
            Height          =   330
            Left            =   570
            TabIndex        =   11
            Top             =   285
            Width           =   1545
            _ExtentX        =   2249
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
            Object.Width           =   1545
            cSize           =   1
            cChkResult      =   0   'False
            cCalenderPosition=   1
            cType           =   1
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
         Begin Cis3D_v60.CIS3D CIS3D20 
            Height          =   255
            Left            =   480
            Top             =   30
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   450
            ForeColor       =   16711680
            Caption         =   "有 効 期 限"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   2
            cPositionY      =   -15
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   330
            Left            =   60
            Top             =   285
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   582
            ForeColor       =   0
            Caption         =   "開始"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   330
            Left            =   60
            Top             =   630
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   582
            ForeColor       =   0
            Caption         =   "終了"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
      End
      Begin Cis3D_v60.CIS3D K_Sykbnm2 
         Height          =   1005
         Left            =   90
         Top             =   90
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   1773
         ForeColor       =   16711680
         Caption         =   "追加"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   120
      End
      Begin Cis3D_v60.CIS3D B1LB_Back 
         Height          =   1005
         Left            =   540
         Top             =   90
         Width           =   6150
         _ExtentX        =   10848
         _ExtentY        =   1773
         BackColor       =   14737632
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
         cBoderColor1    =   12632256
         cBoderColor2    =   12632256
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cBoderStyle     =   2
         Begin Cis3D_v60.CIS3D BLB_Zaist 
            Height          =   495
            Left            =   0
            Top             =   0
            Width           =   4230
            _ExtentX        =   7461
            _ExtentY        =   873
            ForeColor       =   16711680
            Caption         =   "材 質"
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
            cAlingnment     =   1
            cPositionX      =   90
            Begin CisText_V60.CisText iB1_Zaist 
               Height          =   360
               Left            =   960
               TabIndex        =   1
               Top             =   75
               Width           =   3180
               _ExtentX        =   5609
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
               cAutoEject      =   0   'False
            End
         End
         Begin Cis3D_v60.CIS3D BLB_Size 
            Height          =   495
            Left            =   4230
            Top             =   0
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   873
            ForeColor       =   16711680
            Caption         =   "寸 法"
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
            cAlingnment     =   1
            cPositionX      =   90
            Begin Cis3D_v60.CIS3D B1lb_Size 
               Height          =   360
               Left            =   960
               Top             =   75
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   635
               BackColor       =   16777152
               Caption         =   "XXXXX"
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
         Begin Cis3D_v60.CIS3D BLB_SizeBack 
            Height          =   945
            Left            =   0
            Top             =   510
            Width           =   6150
            _ExtentX        =   10848
            _ExtentY        =   1667
            BackColor       =   16777152
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
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   2
            Begin Cis3D_v60.CIS3D B1LB_Itatu 
               Height          =   495
               Left            =   0
               Top             =   0
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   873
               ForeColor       =   16711680
               Caption         =   "板厚"
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
               cAlingnment     =   1
               cPositionX      =   90
               Begin CisText_V60.CisText iB1_Itatu 
                  Height          =   360
                  Left            =   600
                  TabIndex        =   2
                  Top             =   75
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   635
                  cFormat         =   "###0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###0.00"
                  MaxLength       =   7
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "###0.00"
                  cILength        =   4
               End
            End
            Begin Cis3D_v60.CIS3D B1LB_Width 
               Height          =   495
               Left            =   1605
               Top             =   0
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   873
               ForeColor       =   16711680
               Caption         =   " 幅"
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
               cAlingnment     =   1
               cPositionX      =   90
               Begin CisText_V60.CisText iB1_Width 
                  Height          =   360
                  Left            =   600
                  TabIndex        =   3
                  Top             =   75
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   635
                  cFormat         =   "###0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###0.00"
                  MaxLength       =   7
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "###0.00"
                  cILength        =   4
               End
            End
            Begin Cis3D_v60.CIS3D B1LB_Long 
               Height          =   495
               Left            =   3210
               Top             =   0
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   873
               ForeColor       =   16711680
               Caption         =   "長さ"
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
               cAlingnment     =   1
               cPositionX      =   90
               Begin CisText_V60.CisText iB1_Long 
                  Height          =   360
                  Left            =   600
                  TabIndex        =   4
                  Top             =   75
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   635
                  cFormat         =   "###0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###0.00"
                  MaxLength       =   7
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "###0.00"
                  cILength        =   4
               End
            End
            Begin Cis3D_v60.CIS3D B1LB_BackTani 
               Height          =   495
               Left            =   4815
               Top             =   0
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   873
               ForeColor       =   16711680
               Caption         =   "単位"
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
               cAlingnment     =   1
               cPositionX      =   90
               Begin Cis3D_v60.CIS3D B1lb_Tani 
                  Height          =   360
                  Left            =   630
                  Top             =   75
                  Width           =   585
                  _ExtentX        =   1032
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "XXXX"
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
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   1005
         Left            =   6690
         Top             =   90
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1773
         ForeColor       =   16711680
         Caption         =   "ピッチ"
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
         Begin CisText_V60.CisText iB1_Pitch 
            Height          =   360
            Left            =   30
            TabIndex        =   5
            Top             =   465
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            cFormat         =   "###0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "###0.00"
            MaxLength       =   7
            cDataType       =   2
            cDataReplace    =   1
            cFcDecimal      =   2
            cFdAutoFormat   =   1
            cGFormat        =   "###0.00"
            cILength        =   3
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   1005
         Left            =   7725
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   1773
         ForeColor       =   16711680
         Caption         =   "取 数"
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
         Begin CisText_V60.CisText iB1_TSur 
            Height          =   360
            Left            =   45
            TabIndex        =   6
            Top             =   465
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   635
            cFormat         =   "####"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "####"
            MaxLength       =   4
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "####"
            cILength        =   4
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D19 
         Height          =   1005
         Left            =   8490
         Top             =   90
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   1773
         ForeColor       =   16711680
         Caption         =   "所要量"
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
         Begin CisText_V60.CisText iB1_Siyou 
            Height          =   360
            Left            =   30
            TabIndex        =   7
            Top             =   465
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   635
            cFormat         =   "###0.0000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "###0.0000"
            MaxLength       =   9
            cDataType       =   2
            cDataReplace    =   1
            cFbComma        =   0
            cFcDecimal      =   4
            cFdAutoFormat   =   1
            cGFormat        =   "###0.0000"
            cILength        =   4
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1005
         Left            =   9900
         Top             =   90
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   1773
         ForeColor       =   16711680
         Caption         =   "ｽｸﾗｯﾌﾟ重量"
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
         Begin CisText_V60.CisText iB1_MSiyou 
            Height          =   360
            Left            =   75
            TabIndex        =   8
            Top             =   465
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            cFormat         =   "####0.000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "####0.000"
            MaxLength       =   9
            cDataType       =   2
            cDataReplace    =   1
            cFbComma        =   0
            cFcDecimal      =   3
            cFdAutoFormat   =   1
            cGFormat        =   "####0.000"
            cILength        =   4
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1005
         Left            =   11250
         Top             =   90
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   1773
         ForeColor       =   16711680
         Caption         =   "展開区"
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
         Begin VB.CheckBox B1Chk_Tenkb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "有"
            Height          =   345
            Left            =   255
            TabIndex        =   9
            Top             =   480
            Width           =   555
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   1005
         Left            =   12120
         Top             =   90
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   1773
         ForeColor       =   16711680
         Caption         =   "ＬＴ"
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
         Begin CisText_V60.CisText iB1_LT 
            Height          =   360
            Left            =   120
            TabIndex        =   10
            Top             =   465
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            cFormat         =   "##"
            Enabled         =   0   'False
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
      Height          =   975
      Left            =   450
      ScaleHeight     =   915
      ScaleWidth      =   8655
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1020
      Width           =   8715
      Begin Cis3D_v60.CIS3D HLB_Hinbn 
         Height          =   795
         Left            =   90
         Top             =   60
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "品    番"
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
            TabIndex        =   0
            Top             =   360
            Width           =   3210
            _ExtentX        =   5662
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
      Begin Cis3D_v60.CIS3D HLB_Seban 
         Height          =   795
         Left            =   3450
         Top             =   60
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   1402
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
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D H1lb_Seban 
            Height          =   360
            Left            =   120
            Top             =   360
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXX"
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
      Begin Cis3D_v60.CIS3D HLB_Hinnm 
         Height          =   795
         Left            =   4590
         Top             =   60
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "品      名"
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
         Begin Cis3D_v60.CIS3D H1lb_Hinnm 
            Height          =   360
            Left            =   120
            Top             =   360
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX1XXXXXXXXX1"
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
   End
   Begin VB.Timer Timer3 
      Left            =   1980
      Top             =   810
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
      Height          =   1365
      Left            =   10470
      ScaleHeight     =   1305
      ScaleWidth      =   4365
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   630
      Width           =   4425
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1755
         Top             =   660
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         Left            =   2625
         Top             =   660
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         Left            =   3480
         Top             =   660
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         Left            =   2610
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         Left            =   900
         Top             =   660
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         Top             =   660
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   1740
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   585
         Left            =   30
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
      Begin CisBtn_60.CisBtn PB_MNT 
         Height          =   585
         Left            =   885
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
      Begin CisBtn_60.CisBtn PB_CHG 
         Height          =   585
         Left            =   3480
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   7
         BtnCaption      =   "代用"
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
         PFCaption       =   "( F6 )"
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
         PFBackColor     =   0
      End
   End
   Begin VB.TextBox Dummy 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   990
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   435
      Width           =   195
   End
   Begin VB.Timer Timer1 
      Left            =   930
      Top             =   660
   End
   Begin VB.Timer Timer2 
      Left            =   1440
      Top             =   720
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
      Caption         =   "【 材料・品番関連付け  】"
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
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'なし
         Height          =   345
         Index           =   1
         Left            =   10620
         ScaleHeight     =   345
         ScaleWidth      =   3675
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   30
         Width           =   3675
         Begin Cis3D_v60.CIS3D CIS3D15 
            Height          =   345
            Left            =   2460
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   8421631
            ForeColor       =   0
            Caption         =   "削除"
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
            cPositionX      =   20
            Begin Cis3D_v60.CIS3D iT1_Delsu 
               Height          =   255
               Left            =   540
               Top             =   30
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   8421631
               ForeColor       =   0
               Caption         =   "ZZ9"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D6 
            Height          =   345
            Left            =   0
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   8454143
            ForeColor       =   0
            Caption         =   "追加"
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
            cPositionX      =   20
            Begin Cis3D_v60.CIS3D iT1_Addsu 
               Height          =   255
               Left            =   540
               Top             =   30
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   8454143
               ForeColor       =   0
               Caption         =   "ZZ9"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D7 
            Height          =   345
            Left            =   1230
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   16777088
            ForeColor       =   0
            Caption         =   "修正"
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
            cPositionX      =   20
            Begin Cis3D_v60.CIS3D iT1_Mntsu 
               Height          =   255
               Left            =   540
               Top             =   30
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   16777088
               ForeColor       =   0
               Caption         =   "ZZ9"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
         End
      End
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14400
         Top             =   90
         Width           =   885
         _ExtentX        =   1561
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
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   285
         Left            =   8070
         Top             =   60
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
End
Attribute VB_Name = "CXM0110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************************'
'**
'**       フォーム名    :   品番・材料関連付け
'**       フォームID    :   CXM0110
'**       処理概要      :
'**
'**       作  成  日    :   2003/11/25  By CIS SIMPR-A
'**       変  更  日    :   2004/04/14  バグ修正
'**       変  更  日    :   2004/11/16  By CIS グリッド内カラム移動
'**       変  更  日    :   2008/02/13  By CIS 千代田工業用カスタマイズ
'**       変  更  日    :   2008/07/15  By CIS 使用量/スクラップ　小数点第３位
'**       変  更  日    :   2008/07/24  By CIS 使用量 自動計算中止
'**       変  更  日    :   2009/09/22  By CIS 土岐津ver
'**       変  更  日    :   2010/09/17  By CIS 所要量、小数点以下桁数変更
'**
'**
'**     Private Sub XXXXX()
'**
'**         RV_Left = 0
'**         RV_Top = 0
'**         RV_Hinbn = RTrim(iH1_Hinbn)
'**         RV_Call2 = "HIN"
'**
'**         CXM0110.Show vbModal
'**         Unload CXM0110
'**         Set CXM0110 = Nothing
'**
'**    End Sub
'**
'**    RV_Call2 =  "Hin": 品番 → 材料
'**                "Zai": 材料 → 品番
'**                "   ": Command  = " ":品番 → 材料
'**                                = "1":材料 → 品番
'**
'**
'********************************************************************************************'
    
    Const i_Dragico     As String = "BMP\DRAG1PG.ICO"
    
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)
    Dim SyoriKB2        As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM2        As String       ' 処理名称(新規･修正･削除 等)
    Dim wMNTKb          As String
    
    Dim mCHK            As Boolean
    Dim RowColFlg       As Boolean
    Dim GridBack        As String
    
    Dim DayString()     As String
    
    Dim mEdit           As String

    Dim mGotColor       As Long
    Dim mLostColor      As Long
    
    Dim CisVsGrid1      As New CisVsGrid3
        
    Dim mIx             As Long
    Dim mIx2            As Long
    Dim mStr            As String
    
    Dim ActObj          As Object
    Dim ActRow          As Integer
    
    Dim KCol            As Long         ' 非表示項目開始ｶﾗﾑ
    
    Dim SCol()          As Long         ' 表示項目ｶﾗﾑ(品番)
    Dim SColZ()         As Long         ' 表示項目ｶﾗﾑ(材料)
    Dim SColS()         As Long         ' 表示項目ｶﾗﾑ(材料)
    
    Dim wKanri          As Long         ' 材料管理番号
    
    Dim wsZAISY         As String       ' 材種
    Dim wsKeijo         As String       ' 形状
    Dim wsZaist         As String       ' 材質
    Dim wsSize          As Byte         ' 寸法区分

    Dim wBK_Row          As Long
    Dim wBk_Kanri        As Long
    Dim wBk_Hinbn        As String
    
    Dim wsKari          As Byte
    
    Dim Daiyo_Flg       As Boolean
    Dim CisVsGrid2      As New CisVsGrid3
    Dim D_ZKanriNo      As Long
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()

    If Rv_Call2 = "" Then
'   #------------------#
'   # 二 重 起 動 防 止 #
'   #------------------#
        If Not Dupli_Start Then End
'   #------------------#
'   # 初 期 内 容 取 得 #
'   #------------------#
        If Not IniGet Then End
    End If

'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Set CisFrm = New CISFormContorl
'
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
'
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
'
    Dummy.Left = -1000
    Dummy.Enabled = False
'    Dummy2.Left = -1000
'    Dummy3.Left = -1000
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔(基本)
    Timer2.Enabled = False   ' 処理区分点滅禁止(基本)
'    Timer3.Interval = 500    ' 処理区分点滅間隔(取引先)
'    Timer3.Enabled = False   ' 処理区分点滅禁止(取引先)
'
    wMNTKb = Rv_Call2
    If RTrim(wMNTKb) = "" Then
        If RTrim(Command) = "" Then
            wMNTKb = "HIN"
        Else
            wMNTKb = "ZAI"
        End If
    End If
    
    '+--------------------------------+
    '+  グリッド初期設定(材料⇒品番)
    '+--------------------------------+
    With CisVsGrid1
        Set .GridObj = VsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 350
        .PatanMax = 1
        .SelectCol = 22
        .InitGet = False
        .Init
    End With
    Call GridInit1

'   [ Drag Icon Set ]
    On Error Resume Next
    Set VsGrid1.DragIcon = LoadPicture(gIniExe & i_Dragico)
    On Error GoTo 0

    Back_Daiyo.Visible = False
    Daiyo_Flg = False
    
    '+------------------------------------+
    '+  代用寸法入力グリッド初期設定
    '+------------------------------------+
    With CisVsGrid2
        Set .GridObj = VsGrid2
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 5
        .InitGet = False
        .Init
    End With
    Call GridInit2


'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = "ADD"
    Call SyoriName(SyoriKB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)
'
    If Rv_Call2 = "" Then
        Me.Show
    End If
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    If wMNTKb = "HIN" Then
        PGName = "【品番・材料関連付け】"
        HLB_Hinbn = "品    番"
        HLB_Seban = "背番号"
        HLB_Hinnm = "品    名"
        BLB_Zaist = "材 質"
        BLB_Size = "寸 法"
    Else
        PGName = "【材料・品番関連付け】"
        HLB_Hinbn = "材    質"
        HLB_Seban = "寸  法"
        HLB_Hinnm = "材 料 名"
        BLB_Zaist = "品 番"
        BLB_Size = "背番号"
    End If

    wBk_Kanri = 0
    wBk_Hinbn = ""
    If Rv_Call2 <> "" Then: Call Disp_Call
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
           Case vbKeyF1:        If PB_ADD.Visible Then Call PB_ADD_Click                        '【新規】
           Case vbKeyF2:        If PB_MNT.Visible Then Call PB_MNT_Click                        '【修正】
           Case vbKeyF3:        If PB_DEL.Visible Then Call PB_DEL_Click                        '【削除】
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click                      '【検索】
           Case vbKeyPageUp:    If PB_BACK.Visible Then Call PB_Back_Click Else Exit Function   '【後退】
           Case vbKeyPageDown:  If PB_NEXT.Visible Then Call PB_Next_Click Else Exit Function   '【前進】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click                        '【取消】
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click                        '【終了】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                                          '【入力】
                                    If ProcHB = "B2" Then
                                        Call vsGrid1_DblClick
                                    Else
                                        Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                        Call PB_ENT_Click
                                        Dummy.Enabled = False
                                    End If
                               End If
           Case vbKeyF6:       If PB_CHG.Visible Then Call PB_CHG_Click    '【代用】代用寸法入力
           Case Else:          Exit Function
    End Select
    Key_Acc = True
End Function
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Rv_Call2 = "" Then Cancel = 1
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      呼出された場合の初期表示
'+----------------------------------------------------------+
Private Sub Disp_Call()
    If wMNTKb = "HIN" Then
        iH1_Hinbn = RV_Hinbn
        Call iH1_Hinbn_LostFocus
    Else
        iH1_Hinbn = RV_Zaist
        Call iH1_Hinbn_LostFocus
        Call Get_Size(1)
        
        BLB_SizeBack.Enabled = False        ' 板厚、幅、長さ　入力不可
    End If
    
    gSL_Select = "SELECT * "
    gSL_Select = gSL_Select & " FROM 材料品番マスタ ZH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料マスタ ZR "
    gSL_Select = gSL_Select & "   ON ZH.材料管理番号 = ZR.材料管理番号 "
    If wMNTKb = "HIN" Then
        gSL_Select = gSL_Select & " WHERE ZH.使用品番 = '" & RTrim(iH1_Hinbn) & "'"
    Else
        gSL_Select = gSL_Select & " WHERE ZR.材質 = '" & RTrim(iH1_Hinbn) & "'"
    End If
    If Not DBInput("INV") Then
        SyoriKB = "ADD"
    Else
        SyoriKB = "MNT"
    End If
    Call SyoriName(SyoriKB)
    Call ReturnPress
    
    If wMNTKb = "ZAI" Then: Call Disp_Call_Size
End Sub
Private Sub Disp_Call_Size()
    If Rv_Call2 = "" Then Exit Sub
    
    iB1_Itatu = CisFun.RSetFld(RV_Itatu, 7, "###0.00")
    iB1_Width = CisFun.RSetFld(RV_Width, 7, "###0.00")
    iB1_Long = CisFun.RSetFld(RV_Long, 7, "###0.00")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT START
    B1lb_Tani = RTrim(RV_Tani)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT END
End Sub

'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    If Rv_Call2 = "" Then
    '+---------------------+
    '+   ﾃﾞｰﾀﾍﾞｰｽ切断
    '+---------------------+
        Call CisDB.DBDISConnect
    '+---------------------+
    '+   ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
    '+---------------------+
        Call CisFrm.UnLoadDisp
        End
    Else
        Unload CXM0110
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
Private Sub PB_CAN_Click()
    Call CisFrm.MousePT(11)
   
    If ProcHB = "B1" Then
        ProcHB = "H1"
        If SyoriKB2 <> "ADD" Then
            ProcHB = "B2"
            Call DispChange(ProcHB)
            Call HeadBodyClear("B1")
            Call CisVsGrid1.vsColor(消去個別)
        Else
            If Rv_Call2 <> "" Then: Call PB_END_Click
            Call DispChange(ProcHB)
            Call HeadBodyClear("B")
        End If
        GoTo PB_CAN_Ed
    End If
            
    If ProcHB = "B2" Then
        If Rv_Call2 <> "" Then: Call PB_END_Click
        ProcHB = "H1"
        Call DispChange(ProcHB)
        SyoriKB2 = ""
        Call SyoriName2(SyoriKB2)
        Call HeadBodyClear("B")
        Call SyoriName(SyoriKB)
        GoTo PB_CAN_Ed
    End If

PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    新　規  キ ー(F1)        +
'+----------------------------+
Private Sub PB_ADD_Click()
    
    If ProcHB = "H1" Then
        If SyoriKB <> "ADD" Then
            SyoriKB = "ADD"
            Call SyoriName(SyoriKB)
        End If
    Else
        If SyoriKB2 <> "ADD" Then
            ProcHB = "B1"
            Call DispChange(ProcHB)
            SyoriKB2 = "ADD"
            Call SyoriName2(SyoriKB2)
            Call HeadBodyClear("B1")
            Call CisVsGrid1.vsColor(消去個別)
            iB1_Zaist.SetFocus
        End If
    End If
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)        +
'+----------------------------+
Private Sub PB_MNT_Click()
    
    If ProcHB = "H1" Then
        If SyoriKB <> "MNT" Then
            SyoriKB = "MNT"
            Call SyoriName(SyoriKB)
        End If
    Else
        If SyoriKB2 <> "MNT" Then
            ProcHB = "B2"
            Call DispChange(ProcHB)
            SyoriKB2 = "MNT"
            Call SyoriName2(SyoriKB2)
            Call CisVsGrid1.vsColor(消去個別)
            VsGrid1.SetFocus
        
            If Trim(VsGrid1.TextMatrix(VsGrid1.Row, SColZ(14))) <> "" Then        ' 14:編集前品番
               Call GridToText_Zai
            End If
        End If
    End If
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)        +
'+----------------------------+
Private Sub PB_DEL_Click()
    If ProcHB = "H1" Then
        If SyoriKB <> "DEL" Then
            SyoriKB = "DEL"
            Call SyoriName(SyoriKB)
        End If
    Else
        If SyoriKB2 <> "DEL" Then
            
            ProcHB = "B2"
            SyoriKB2 = "DEL"
            Call DispChange(ProcHB)
            Call SyoriName2(SyoriKB2)
            Call CisVsGrid1.vsColor(消去個別)
            VsGrid1.SetFocus
            
            If Trim(VsGrid1.TextMatrix(VsGrid1.Row, SColZ(14))) <> "" Then        ' 14:編集前品番
               Call GridToText_Zai
            End If
        End If
    End If

End Sub
'+----------------------------+
'+    代　用 キ ー(F6)       +
'+----------------------------+
Private Sub PB_CHG_Click()
    If Not Daiyo_Flg Then
       ' 材料管理番号　検索
       If wMNTKb = "HIN" Then
       Else
          If RTrim(iB1_Zaist) = "" Then
              With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    品番を入力して下さい          "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
              End With
              iB1_Zaist.SetFocus
              Exit Sub
          End If
          gSL_Select = "SELECT 品番 FROM 品番マスタ "
          gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iB1_Zaist) & "'"
          If Not HNMRead(gSL_Select, 1) Then
              With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    品番マスタ 未登録              "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
              End With
              Call FocusIchi_Set        ' 2009.09.24 upd : iB1_Itatu.SetFocus
              Exit Sub
          End If
       End If
       
       gSL_Select = "SELECT 材料管理番号 FROM 材料マスタ"
       If wMNTKb = "HIN" Then
           gSL_Select = gSL_Select & " WHERE 材質 = '" & RTrim(iB1_Zaist) & "'"
       Else
           gSL_Select = gSL_Select & " WHERE 材質 = '" & RTrim(iH1_Hinbn) & "'"
       End If
       gSL_Select = gSL_Select & "   AND 板厚 = " & CisFun.Val2(iB1_Itatu)
       gSL_Select = gSL_Select & "   AND 幅 = " & CisFun.Val2(iB1_Width)
       gSL_Select = gSL_Select & "   AND 長さ = " & CisFun.Val2(iB1_Long)
       If Not ZRMRead(gSL_Select, 1) Then
           With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    材料マスタ 未登録              "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
           End With
          Call FocusIchi_Set            ' 2009.09.24 upd : iB1_Itatu.SetFocus
          Exit Sub
       End If
       D_ZKanriNo = ZRM.材料管理番号
       
       ' 材料品番代用マスタ情報
       gSL_Select = "Select * from 材料代用マスタ "
       gSL_Select = gSL_Select & " Where 材料管理番号 = " & D_ZKanriNo
       If wMNTKb = "HIN" Then
           gSL_Select = gSL_Select & " AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
       Else
           gSL_Select = gSL_Select & " AND 品番 = '" & RTrim(iB1_Zaist) & "'"
       End If
       If ZDARead(gSL_Select, 0, 1) Then
          Call GridSet2
       End If
       Call ZDAClose(1)
       
       VsGrid2.Row = 1
       Call VsGridToText2(1)
       
       Call TextToVsGrid2(1)
       
       Back_Daiyo.Visible = True
       Daiyo_Flg = True
       PB_CAN.Visible = False
       PB_ENT.Visible = False
       PB_CHG.BtnCaption = "終了"
       iB1_DItatu.SetFocus
       B1_Area1.Enabled = False
       B2_Area1.Enabled = False
    
    Else
       Call TextToVsGrid2(Val(iB1_DItatu.Tag))
       If Not Grid2Chk Then Exit Sub            '入力内容ﾁｪｯｸ
       Call VsGrid2_DBInsert                    '登録内容DB更新

       Back_Daiyo.Visible = False
       Daiyo_Flg = False
       PB_CAN.Visible = True
       PB_ENT.Visible = True
       PB_CHG.BtnCaption = "代用"
    
       CisVsGrid2.Clear
       B1_Area1.Enabled = True
       B2_Area1.Enabled = True
       
       Call FocusIchi_Set           ' 2009.09.24 upd : iB1_Itatu.SetFocus
       
    End If
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    
    If wMNTKb = "HIN" Then
        gSL_Select = " SELECT ZH.使用品番,HN.社内背番 背番号 "
        gSL_Select = gSL_Select & " FROM 材料品番マスタ ZH "
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
        gSL_Select = gSL_Select & "   ON HN.品番 = ZH.使用品番 "
        gSL_Select = gSL_Select & " WHERE ZH.使用品番 < '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " GROUP BY ZH.使用品番,HN.社内背番"
        gSL_Select = gSL_Select & " ORDER BY ZH.使用品番 Desc "
    Else
        gSL_Select = " SELECT ZR.材質"
        gSL_Select = gSL_Select & " FROM 材料マスタ ZR"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料品番マスタ ZH"
        gSL_Select = gSL_Select & "   ON ZR.材料管理番号 = ZH.材料管理番号"
        gSL_Select = gSL_Select & " Where ZH.使用品番 Is Not Null"
        gSL_Select = gSL_Select & "   AND ZR.材質 < '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " GROUP BY ZR.材質"
        gSL_Select = gSL_Select & " ORDER BY ZR.材質 DESC"
    End If
    
    Call DBInput("Back")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Hinbn.SetFocus
    Dummy.Enabled = False
    
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    
    If wMNTKb = "HIN" Then
        gSL_Select = " SELECT ZH.使用品番,HN.社内背番 背番号"
        gSL_Select = gSL_Select & " FROM 材料品番マスタ ZH "
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
        gSL_Select = gSL_Select & "   ON HN.品番 = ZH.使用品番 "
        gSL_Select = gSL_Select & " WHERE ZH.使用品番 > '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " GROUP BY ZH.使用品番,HN.社内背番"
        gSL_Select = gSL_Select & " ORDER BY ZH.使用品番 "
    Else
        gSL_Select = " SELECT ZR.材質"
        gSL_Select = gSL_Select & " FROM 材料マスタ ZR"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料品番マスタ ZH"
        gSL_Select = gSL_Select & "   ON ZR.材料管理番号 = ZH.材料管理番号"
        gSL_Select = gSL_Select & " Where ZH.使用品番 Is Not Null"
        gSL_Select = gSL_Select & "   AND ZR.材質 > '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " GROUP BY ZR.材質 "
        gSL_Select = gSL_Select & " ORDER BY ZR.材質 "
    End If
    
    
    Call DBInput("Next")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Hinbn.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iB1_Itatu" Then Call Look_Zairy: Exit Sub
    If PB_Look.Tag = "iB1_Witdh" Then Call Look_Zairy: Exit Sub
    If PB_Look.Tag = "iB1_Long" Then Call Look_Zairy: Exit Sub
    If wMNTKb = "HIN" Then
        If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
        If PB_Look.Tag = "iB1_Zaist" Then Call Look_Zairy: Exit Sub
    Else
        If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Zaist: Exit Sub
        If PB_Look.Tag = "iB1_Zaist" Then Call Look_Hinbn: Exit Sub
    End If
'
    If PB_Look.Tag = "iB1_SDate" Then
        iB1_SDate.ShowCalender
        Exit Sub
    End If
    If PB_Look.Tag = "iB1_EDate" Then
        iB1_EDate.ShowCalender
        Exit Sub
    End If

End Sub
'+----------------------------+
'+       品番検索              +
'+----------------------------+
Private Sub Look_Hinbn()

    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RVI_Hinmoku = 9
    
    CKK0030.Show vbModal
    Unload CKK0030
    Set CKK0030 = Nothing

    If RV_Rtn Then
        If wMNTKb = "HIN" Then
            iH1_Hinbn = RTrim(RV_Hinbn)
            H1lb_Seban = RTrim(RV_SSeban)
            H1lb_Hinnm = RTrim(RV_Hinnm)
        Else
            iB1_Zaist = RTrim(RV_Hinbn)
            B1lb_Size = RTrim(RV_SSeban)
            
        End If
    End If
End Sub
'+----------------------------+
'+       材質検索              +
'+----------------------------+
Private Sub Look_Zaist()

    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    
    CKK0075.Show vbModal
    Unload CKK0075
    Set CKK0075 = Nothing

    If RV_Rtn Then
        If wMNTKb = "HIN" Then
            iB1_Zaist = RTrim(RV_Zaist)
            Call Get_Size(0)
        Else
            iH1_Hinbn = RTrim(RV_Zaist)
            Call Get_Size(1)
            H1lb_Hinnm = RTrim(RV_Zainm)
        End If
    End If
End Sub
'+----------------------------+
'+       材料(サイズ)検索      +
'+----------------------------+
Private Sub Look_Zairy()

    RV_Left = 0
    RV_Top = 0
    If wMNTKb = "HIN" Then
        RV_Zaist = RTrim(iB1_Zaist)
    Else
        RV_Zaist = RTrim(iH1_Hinbn)
    End If
    RV_Call = "C"

    CKK0070.Show vbModal
    Unload CKK0070
    Set CKK0070 = Nothing

    If RV_Rtn Then
        If wMNTKb = "HIN" Then: iB1_Zaist = RTrim(RV_Zaist)
        Call Get_Size(0)
        
        If RV_Itatu <> 0 Then: iB1_Itatu = CisFun.RSetFld(RV_Itatu, 7, "###0.00")
        If RV_Width <> 0 Then: iB1_Width = CisFun.RSetFld(RV_Width, 7, "###0.00")
        If RV_Long <> 0 Then: iB1_Long = CisFun.RSetFld(RV_Long, 7, "###0.00")
        If RV_Tani <> "" Then: B1lb_Tani = RV_Tani
    End If
End Sub
'+---( H1 )----------------------------------------------------------------------------
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
    
    H1lb_Seban = ""
    H1lb_Hinnm = ""
    If wMNTKb = "HIN" Then
        gSL_Select = "SELECT * FROM 品番マスタ "
        gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
        Call HNMRead(gSL_Select, 1)
        H1lb_Seban = RTrim(HNM.社内背番)
        H1lb_Hinnm = RTrim(HNM.品名)
    Else
        gSL_Select = "SELECT * FROM 材料マスタ "
        gSL_Select = gSL_Select & " WHERE 材質 = '" & RTrim(iH1_Hinbn) & "'"
        Call ZRMRead(gSL_Select, 1)
        If ZRM.寸法区分 <> 0 Then: H1lb_Seban = "無"
        Call Get_Size(0)
        H1lb_Hinnm = RTrim(ZRM.材料名)
    End If
End Sub
'+---( B1 )----------------------------------------------------------------------------
'+-----------------------------+
'+  iB1_Zaist
'+-----------------------------+
Private Sub iB1_Zaist_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Zaist_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B1lb_Size = ""
    If wMNTKb = "HIN" Then
        gSL_Select = "SELECT * FROM 材料マスタ "
        gSL_Select = gSL_Select & " WHERE 材質 = '" & RTrim(iB1_Zaist) & "'"
        Call ZRMRead(gSL_Select, 1)
        If ZRM.寸法区分 <> 0 Then: B1lb_Size = "無"
        
        Call Get_Size(0)
    Else
        gSL_Select = "SELECT * FROM 品番マスタ "
        gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iB1_Zaist) & "'"
        Call HNMRead(gSL_Select, 1)
        B1lb_Size = RTrim(HNM.社内背番)
    End If
End Sub
'+-----------------------------+
'+  iB1_Itatu
'+-----------------------------+
Private Sub iB1_Itatu_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Itatu_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT START
    If wMNTKb = "HIN" Then
       If RTrim(iB1_Zaist) <> "" Then
          Call Disp_Tani(RTrim(iB1_Zaist))
       End If
    Else
       If RTrim(iH1_Hinbn) <> "" Then
          Call Disp_Tani(RTrim(iH1_Hinbn))
       End If
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT END
End Sub
'+-----------------------------+
'+  iB1_Width
'+-----------------------------+
Private Sub iB1_Width_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Width_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT START
    If wMNTKb = "HIN" Then
       If RTrim(iB1_Zaist) <> "" Then
          Call Disp_Tani(RTrim(iB1_Zaist))
       End If
    Else
       If RTrim(iH1_Hinbn) <> "" Then
          Call Disp_Tani(RTrim(iH1_Hinbn))
       End If
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT END
End Sub
'+-----------------------------+
'+  iB1_Long
'+-----------------------------+
Private Sub iB1_Long_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Long_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT START
    If wMNTKb = "HIN" Then
       If RTrim(iB1_Zaist) <> "" Then
          Call Disp_Tani(RTrim(iB1_Zaist))
       End If
    Else
       If RTrim(iH1_Hinbn) <> "" Then
          Call Disp_Tani(RTrim(iH1_Hinbn))
       End If
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT END
End Sub
'+-----------------------------+
'+  iB1_Pitch
'+-----------------------------+
Private Sub iB1_Pitch_LostFocus()
'=============================== *** 2008/07/24 Start
'    If CisFun.Val2(iB1_Pitch) <> 0 And _
'       CisFun.Val2(iB1_TSur) <> 0 Then
'        iB1_Siyou = CisFun.RSetFld(CisFun.HasuProc(CisFun.Val2(iB1_Pitch) / CisFun.Val2(iB1_TSur), 切上げ, 2), 7, "###0.00")
'    End If
'=============================== *** 2008/07/24 End
End Sub
'+-----------------------------+
'+  iB1_TSur
'+-----------------------------+
Private Sub iB1_TSur_LostFocus()
'=============================== *** 2008/07/24 Start
'    If CisFun.Val2(iB1_Pitch) <> 0 And _
'       CisFun.Val2(iB1_TSur) <> 0 Then
'        iB1_Siyou = CisFun.RSetFld(CisFun.HasuProc(CisFun.Val2(iB1_Pitch) / CisFun.Val2(iB1_TSur), 切上げ, 2), 7, "###0.00")
'    End If
'=============================== *** 2008/07/24 End
End Sub
'+-----------------------------+
'+  B1Chk_Tenkb
'+-----------------------------+
Private Sub B1Chk_Tenkb_Click()
    If B1Chk_Tenkb.Value = 0 Then
        B1Chk_Tenkb.Caption = "有"
    Else
        B1Chk_Tenkb.Caption = "無"
    End If
End Sub
'+-----------------------------+
'+  iB1_SDate
'+-----------------------------+
Private Sub iB1_SDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_SDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_EDate
'+-----------------------------+
Private Sub iB1_EDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_EDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+---------------------------------------------+
'+  iB1_DItatu    代用板厚
'+---------------------------------------------+
Private Sub iB1_DItatu_GotFocus()
    Set ActObj = iB1_DItatu
End Sub
Private Sub iB1_DItatu_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange2(KeyCode)
End Sub
Private Sub iB1_DItatu_LostFocus()
    With VsGrid2
        ActRow = CisFun.Val2(iB1_DItatu.Tag)
        If ActRow > .Rows - 1 Then Exit Sub
        .TextMatrix(ActRow, 1) = RTrim(iB1_DItatu)
            
        If Trim(iB1_DItatu) = "" Then
            For gInt = 2 To .Cols - 1
                .TextMatrix(ActRow, gInt) = ""
            Next gInt
            iB1_DWidth = ""
            iB1_DLong = ""
        End If
        
    End With
End Sub
'+---------------------------------------------+
'+  iB1_DWidth    代用幅
'+---------------------------------------------+
Private Sub iB1_DWidth_GotFocus()
    Set ActObj = iB1_DWidth
End Sub
Private Sub iB1_DWidth_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange2(KeyCode)
End Sub
Private Sub iB1_DWidth_LostFocus()
    With VsGrid2
        .TextMatrix(CisFun.Val2(iB1_DItatu.Tag), 2) = iB1_DWidth
    End With
End Sub
'+---------------------------------------------+
'+  iB1_DLong    代用長さ
'+---------------------------------------------+
Private Sub iB1_DLong_GotFocus()
    Set ActObj = iB1_DLong
End Sub
Private Sub iB1_DLong_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange2(KeyCode)
End Sub
Private Sub iB1_DLong_LostFocus()
    With VsGrid2
        .TextMatrix(CisFun.Val2(iB1_DItatu.Tag), 3) = iB1_DLong
    End With
End Sub
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT START
'+------------------------------------+
'+     材料マスタ読込 ＆ 単位表示     +
'+------------------------------------+
Private Sub Disp_Tani(Zaist As String)
    B1lb_Tani = ""

    gSL_Select = ""
    gSL_Select = gSL_Select & " SELECT ISNULL(SY.値名称, '') 単位"
    gSL_Select = gSL_Select & "   FROM 材料マスタ ZR"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 名称マスタ SY"
    gSL_Select = gSL_Select & "             ON SY.区分名称 = '単位区分'"
    gSL_Select = gSL_Select & "            AND SY.値 = ZR.単位"
    gSL_Select = gSL_Select & "            AND SY.区分タイプ = ''"
    gSL_Select = gSL_Select & "  WHERE ZR.材質 = '" & Zaist & "'"
    If CisFun.Val2(iB1_Itatu) <> 0 Then
       gSL_Select = gSL_Select & "    AND ZR.板厚 = " & CisFun.Val2(iB1_Itatu)
    End If
    If CisFun.Val2(iB1_Width) <> 0 Then
       gSL_Select = gSL_Select & "    AND   ZR.幅 = " & CisFun.Val2(iB1_Width)
    End If
    If CisFun.Val2(iB1_Long) <> 0 Then
       gSL_Select = gSL_Select & "    AND ZR.長さ = " & CisFun.Val2(iB1_Long)
    End If
    
    If ZRMRead(gSL_Select, 1) Then
       B1lb_Tani = ZRM.単位
    End If
End Sub
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT END
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
'+------------------------------------------------------------------+
'+      基本情報
'+------------------------------------------------------------------+
Private Sub SyoriName(SYKB As String)
    Dim W_Color     As Long
    W_Color = &H80000008 '初期色ｾｯﾄ(黒)
   
'   H2LB_Size_B.Enabled = False
   
    Select Case SyoriKB
      Case "ADD"
           SyoriNM = "追加"
           K_Sykbnm.ForeColor = &HFFFF&   '(黄色)
           PB_ADD.Visible = False
               
'          H2LB_Size_B.Enabled = True
      Case "MNT"
           SyoriNM = "修正"
           K_Sykbnm.ForeColor = &HFF0000  '(青色)
      
      Case "DEL"
           SyoriNM = "削除"
           K_Sykbnm.ForeColor = &HFF&     '(赤色)
           W_Color = &HFF&
      Case "LOK"
           SyoriNM = "照会"
           K_Sykbnm.ForeColor = &HFF00FF   '(ピンク)
           PB_ADD.Visible = False
           PB_MNT.Visible = False
           PB_DEL.Visible = False
      Case Else
           SyoriNM = ""
    End Select
    For Each gObj In Me.Controls
        If gObj.Name Like "i*" Then gObj.ForeColor = W_Color
        If gObj.Name Like "*Comb*" Then gObj.ForeColor = W_Color
'      * PF-Key 表示設定 *
        If SyoriKB <> "LOC" Then
           If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
              gObj.Name = "PB_DEL" Then
              gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
           End If
        End If
    Next gObj
   
'    If RTrim(PB_Chg.Tag) <> "Hin" Then: PB_DEL.Visible = False
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
'+------------------------------------------------------------------+
'+      詳細情報
'+------------------------------------------------------------------+
Private Sub SyoriName2(SYKB As String)
   Dim W_Color     As Long
   W_Color = &H80000008 '初期色ｾｯﾄ(黒)
   
'   H2LB_Size_B.Enabled = False
    B1LB_Back.Enabled = True
   
   Select Case SyoriKB2
     Case "ADD"
          SyoriNM2 = "追加"
          K_Sykbnm2.ForeColor = &HFFFF&   '(黄色)
          PB_ADD.Visible = False
              
'          H2LB_Size_B.Enabled = True
     Case "MNT"
          SyoriNM2 = "修正"
          K_Sykbnm2.ForeColor = &HFF0000  '(青色)
     
          B1LB_Back.Enabled = False
     Case "DEL"
          SyoriNM2 = "削除"
          K_Sykbnm2.ForeColor = &HFF&     '(赤色)
          W_Color = &HFF&
     Case "LOK"
          SyoriNM2 = "照会"
          K_Sykbnm2.ForeColor = &HFF00FF   '(ピンク)
          PB_ADD.Visible = False
          PB_MNT.Visible = False
          PB_DEL.Visible = False
     Case Else
          SyoriNM2 = ""
   End Select
   For Each gObj In Me.Controls
       If gObj.Name Like "i*" Then gObj.ForeColor = W_Color
       If gObj.Name Like "*Comb*" Then gObj.ForeColor = W_Color
'      * PF-Key 表示設定 *
       If SyoriKB2 <> "LOC" Then
          If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
             gObj.Name = "PB_DEL" Then
             gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB2)
          End If
       End If
   Next gObj
   Timer3.Enabled = False
   If SyoriNM2 = "" Then
      K_Sykbnm2.Visible = False
      K_Sykbnm2.Caption = ""
      Exit Sub
   End If
   K_Sykbnm2.Caption = SyoriNM2
   K_Sykbnm2.Visible = True
   Timer3.Enabled = True
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
'*******************************
'*      処理区分名点滅制御      *
'*******************************
Private Sub Timer3_Timer()
   If K_Sykbnm2.Caption <> "" Then
      K_Sykbnm2.Caption = ""
   Else
      K_Sykbnm2.Caption = SyoriNM2
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
    Dim B2Mode      As Boolean
    Dim B2Color     As Long
    
    H1Mode = False
    H1Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
    B2Mode = False
    B2Color = gPLostSel

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
'          PB_Chg.Tag = "Hin"
'          PB_Chg.BtnCaption = "材料"
          If Rv_Call2 <> "" Then
             PB_END.BtnCaption = "戻る"
          Else
             PB_END.BtnCaption = "終了"
          End If
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
       Case "B2"
          B2Mode = True
          B2Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B2_Area1.Enabled = B2Mode
    B2_Area1.BackColor = B2Color
    

    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    
    If Pro = "B1" Then
       PB_CHG.Visible = True
    Else
       PB_CHG.Visible = False
    End If
    
'    PB_Chg.Visible = H1Mode
'    PB_ZLook.Visible = H1Mode
' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'**************************************
'*      Head/Body 初期化処理           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String
    RowColFlg = True
    
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
                .BackColor = mLostColor
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    
    
    If HeadBody = "B2" Or HeadBody = "B" Or HeadBody = "ALL" Then
        With CisVsGrid1
            Call .Clear
            Call .LayOutChg
                    
'       < カラムワークセット >
            For gLong = 1 To VsGrid1.Cols - 1
                gInt = CisVsGrid1.FixedGetNM(gLong)
                SColZ(gInt) = gLong
            Next gLong
        End With
    End If
    RowColFlg = False
    Exit Sub
HeadBodyClear_ER:
    RowColFlg = False
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        
' << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call HeadBodyClear("B")
            ProcHB = "H1"
            Call DispChange(ProcHB)
            GoTo ReturnPress_Ed
        End If
        
        ProcHB = "B1"
        Call DispChange(ProcHB)
        SyoriKB2 = "ADD"
        Call SyoriName2(SyoriKB2)
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut(0) Then
            GoTo ReturnPress_Ed
        End If
        Call HeadBodyClear("B")
        
        Call GridSet_Zai
        Call Disp_Call_Size
        
        If SyoriKB2 <> "ADD" Then
            ProcHB = "B2"
            Call DispChange(ProcHB)
            VsGrid1.SetFocus
            
            If Trim(VsGrid1.TextMatrix(VsGrid1.Row, SColZ(14))) <> "" Then        ' 14:編集前品番
               Call GridToText_Zai
            End If
        End If
        GoTo ReturnPress_Ed
    End If
        
    If ProcHB = "B2" Then
' << 削除処理 >>
        If SyoriKB2 = "DEL" Then
            If Not DBDeleteM Then
               RowColFlg = True
               Call CisVsGrid1.vsColor(消去個別)
               RowColFlg = False
               GoTo ReturnPress_Ed
            End If
            Call HeadBodyClear("B1")
            Call GridSet_Zai
            Call DispChange(ProcHB)
            Call SyoriName2(SyoriKB2)
            
            If Trim(VsGrid1.TextMatrix(VsGrid1.Row, SColZ(14))) <> "" Then        ' 14:編集前品番
               Call GridToText_Zai
            End If
        Else
' *CHG* 2004/04/14
'            Call vsGrid1_DblClick
'            GridBack = ProcHB
'            ProcHB = "B1"
'            Call DispChange(ProcHB)
        End If
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
' 材種
    If Trim(iH1_Hinbn) = "" Then
        With CisFun
            .MB_Lines = 4
            If wMNTKb = "HIN" Then
                .MB_MSG(2) = "    品番を入力して下さい            "
            Else
                .MB_MSG(2) = "    材種を入力して下さい            "
            End If
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Hinbn.SetFocus
       Exit Function
    End If

    ' マスタチェック
    If wMNTKb = "HIN" Then
        gSL_Select = "SELECT * FROM 品番マスタ "
        gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
        If Not HNMRead(gSL_Select, 1) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    品番マスタ 未登録              "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_Hinbn.SetFocus
           Exit Function
        End If
    Else
        gSL_Select = "SELECT * FROM 材料マスタ "
        gSL_Select = gSL_Select & " WHERE 材質 = '" & RTrim(iH1_Hinbn) & "'"
        If Not ZRMRead(gSL_Select, 1) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    材料マスタ 未登録              "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_Hinbn.SetFocus
           Exit Function
        End If
    End If


'   材料品番マスタ
    gSL_Select = "SELECT * "
    gSL_Select = gSL_Select & " FROM 材料品番マスタ ZH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料マスタ ZR "
    gSL_Select = gSL_Select & "   ON ZH.材料管理番号 = ZR.材料管理番号 "
    If wMNTKb = "HIN" Then
        gSL_Select = gSL_Select & " WHERE ZH.使用品番 = '" & RTrim(iH1_Hinbn) & "'"
    Else
        gSL_Select = gSL_Select & " WHERE ZR.材質 = '" & RTrim(iH1_Hinbn) & "'"
    End If
    If Not DBInput("INV") Then
        If SyoriKB <> "ADD" Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    品番材料関連付け 未登録            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Hinbn.SetFocus
            Exit Function
        End If
    Else
        If SyoriKB = "ADD" Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    品番材料関連付け 登録済            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Hinbn.SetFocus
            Exit Function
        End If
    End If
    
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT START
    If wMNTKb <> "HIN" Then
       Call Disp_Tani(iH1_Hinbn)
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT END
    
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    
    Body1Chk = False
        
    wsKari = 0
    ' 材質
    If RTrim(iB1_Zaist) = "" Then
        With CisFun
            .MB_Lines = 4
            If wMNTKb = "HIN" Then
                .MB_MSG(2) = "    材質を入力して下さい            "
            Else
                .MB_MSG(2) = "    品番を入力して下さい            "
            End If
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_Zaist.SetFocus
       Exit Function
    End If
    
    ' サイズ
    If wMNTKb = "HIN" Then
        gSL_Select = "SELECT * FROM 材料マスタ "
        gSL_Select = gSL_Select & " WHERE 材質 = '" & RTrim(iB1_Zaist) & "'"
        If ZRMRead(gSL_Select, 1) Then
            If ZRM.寸法区分 = 0 Then
                If CisFun.Val2(iB1_Itatu) = 0 And _
                   CisFun.Val2(iB1_Width) = 0 And _
                   CisFun.Val2(iB1_Long) = 0 Then
                     With CisFun
                         .MB_Lines = 5
                         .MB_MSG(2) = "    サイズを入力して下さい          "
                         .MB_Title = ""
                         .MB_Button = Error
                         .MBOX
                     End With
                    Call FocusIchi_Set          ' 2009.09.24 upd : iB1_Itatu.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    
    ' 取数
'=============================== *** 2008/07/24 Start
'    If CisFun.Val2(iB1_TSur) = 0 Then
'        iB1_TSur = CisFun.RSetFld(1, 4, "####")
'    End If
'=============================== *** 2008/07/24 End
    
    ' 使用量
'=============================== *** 2008/07/24 Start
'    If CisFun.Val2(iB1_Pitch) <> 0 And _
'       CisFun.Val2(iB1_TSur) <> 0 Then
'        '=============================================== *** 2008/07/15 Start
'        'iB1_Siyou = CisFun.RSetFld(CisFun.HasuProc(CisFun.Val2(iB1_Pitch) / CisFun.Val2(iB1_TSur), 切上げ, 2), 7, "###0.00")
'        iB1_Siyou = CisFun.RSetFld(CisFun.HasuProc(CisFun.Val2(iB1_Pitch) / CisFun.Val2(iB1_TSur), 切上げ, 2), 7, "###0.000")
'        '=============================================== *** 2008/07/15 End
'    End If
'=============================== *** 2008/07/24 End
    
    ' 使用量
    If CisFun.Val2(iB1_Siyou) = 0 Then
        With CisFun
            .MB_Lines = 5
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 UPDATE START
            '.MB_MSG(2) = "    使用量を入力して下さい          "
'=============================== *** 2008/07/24 Start
'            .MB_MSG(2) = "    単重を入力して下さい          "
            .MB_MSG(2) = "    所要量を入力して下さい        "
'=============================== *** 2008/07/24 End
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 UPDATE END
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_Siyou.SetFocus
       Exit Function
    End If
    
    ' 有効期限
    If RTrim(iB1_SDate) <> "" Then
        If Not iB1_SDate.cChkResult Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    有効期限 開始日エラー              "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iB1_SDate.SetFocus
           Exit Function
        End If
    End If
    
    If RTrim(iB1_EDate) <> "" Then
        If Not iB1_EDate.cChkResult Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    有効期限 終了日エラー              "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iB1_EDate.SetFocus
           Exit Function
        End If
        
        If RTrim(iB1_SDate) > RTrim(iB1_EDate) Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    有効期限エラー              "
                .MB_MSG(4) = "    【開始＞終了】"
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iB1_SDate.SetFocus
           Exit Function
        End If
    End If

    ' 材料マスタ
    Call ItemsClearZRM
    gSL_Select = "SELECT * FROM 材料マスタ "
    If wMNTKb = "HIN" Then
        gSL_Select = gSL_Select & " WHERE 材質 = '" & RTrim(iB1_Zaist) & "'"
    Else
        gSL_Select = gSL_Select & " WHERE 材質 = '" & RTrim(iH1_Hinbn) & "'"
    End If
    gSL_Select = gSL_Select & "   AND 板厚 =  " & CisFun.Val2(iB1_Itatu)
    gSL_Select = gSL_Select & "   AND 幅 =    " & CisFun.Val2(iB1_Width)
    gSL_Select = gSL_Select & "   AND 長さ =  " & CisFun.Val2(iB1_Long)
    If ZRMRead(gSL_Select, 1) Then
        wKanri = ZRM.材料管理番号
        ' 材料品番マスタ
        If SyoriKB2 = "ADD" Then
            gSL_Select = " SELECT ZH.*,ZR.材質,ZR.寸法区分"
            gSL_Select = gSL_Select & " FROM 材料品番マスタ ZH "
            gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料マスタ ZR "
            gSL_Select = gSL_Select & "   ON ZR.材料管理番号 = ZH.材料管理番号 "
            If wMNTKb = "HIN" Then
                gSL_Select = gSL_Select & " WHERE ZH.使用品番 = '" & RTrim(iH1_Hinbn) & "'"
            Else
                gSL_Select = gSL_Select & " WHERE ZH.使用品番 = '" & RTrim(iB1_Zaist) & "'"
            End If
            gSL_Select = gSL_Select & "   AND ZH.材料管理番号 = " & wKanri
            If ZHMRead(gSL_Select, 1) Then
                With CisFun
                    .MB_Lines = 5
                    .MB_MSG(2) = "    品番・材料情報が既に登録されています     "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                iB1_Zaist.SetFocus
                Exit Function
            End If
        End If
    Else
        If wMNTKb = "HIN" Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    材料マスタ 未登録              "
                .MB_MSG(4) = "    材料マスタの仮登録を行いますか？         "
                .MB_Title = ""
                .MB_Button = OK_CAN
                If Not .MBOX Then
                   iB1_Zaist.SetFocus
                   Exit Function
                End If
                wsKari = 1
            End With
        Else
             With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    材料マスタ 未登録              "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
             End With
            Call FocusIchi_Set          ' 2009.09.24 upd : iB1_Itatu.SetFocus
            Exit Function
        End If
    End If
    ' 品番マスタ
    If wMNTKb <> "HIN" Then
        gSL_Select = "SELECT * FROM 品番マスタ "
        gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iB1_Zaist) & "'"
        If Not HNMRead(gSL_Select, 1) Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    品番マスタ 未登録              "
                .MB_MSG(4) = "    品番マスタの仮登録を行いますか？         "
                .MB_Title = ""
                .MB_Button = OK_CAN
                If Not .MBOX Then
                   iB1_Zaist.SetFocus
                   Exit Function
                End If
                wsKari = 1
            End With
        End If
    End If
    
    Body1Chk = True
End Function
'***********************************************************************************
'*
'*       データベース入力 ＆ 表示
'*
'***********************************************************************************
'+---------------------------------+
'+      基本情報
'+---------------------------------+
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    Call HeadBodyClear("B")
    
    If Not ZHMRead(gSL_Select, 1) Then
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
        If wMNTKb = "HIN" Then
            iH1_Hinbn = RTrim(ZHM.使用品番)
            Call iH1_Hinbn_LostFocus
        Else
            iH1_Hinbn = RTrim(ZHM.材質)
            Call iH1_Hinbn_LostFocus
        End If
        
        Call GridSet_Zai
        DBInput = True
    End If

DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'**************************************
'*      寸 法 区 分
'**************************************
Private Sub Get_Size(KB As Byte)
    Static ws_Zaist
    
    If wMNTKb = "HIN" Then
        ws_Zaist = RTrim(iB1_Zaist)
        B1lb_Size = ""
    Else
        ws_Zaist = RTrim(iH1_Hinbn)
        H1lb_Seban = ""
    End If
    wsSize = 0
    
    B1LB_Itatu.Enabled = True
    B1LB_Width.Enabled = True
    B1LB_Long.Enabled = True
    B1LB_BackTani.Enabled = True
    
    If RTrim(ws_Zaist) = "" Then Exit Sub
    
    gSL_Select = "SELECT * FROM 材料マスタ "
    gSL_Select = gSL_Select & " WHERE 材質 = '" & RTrim(ws_Zaist) & "'"
    If Not ZRMRead(gSL_Select, 1) Then Exit Sub
        
    wsSize = ZRM.寸法区分
    If wsSize <> 0 Then
        If wMNTKb = "HIN" Then
            B1lb_Size = "無"
        Else
            H1lb_Seban = "無"
        End If
        B1LB_Itatu.Enabled = False
        B1LB_Width.Enabled = False
        B1LB_Long.Enabled = False
        B1LB_BackTani.Enabled = False
        iB1_Itatu = ""
        iB1_Width = ""
        iB1_Long = ""
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT START
        B1lb_Tani = ""
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT END
    End If

End Sub
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------------+
'+      削 除 処 理            +
'+----------------------------+
Private Function DBDelete()
    DBDelete = False
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
    
    With CisDB
        On Error GoTo DBDelete_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 削除処理 )
' *CHG* 2004/04/14
'        .SQL = "DELETE FROM 材料品番マスタ "
'        .SQL = .SQL & " WHERE 材料管理番号 = " & wKanri
'        .DBExec
        If wMNTKb = "HIN" Then
            .SQL = "DELETE FROM 材料品番マスタ "
            .SQL = .SQL & " WHERE 使用品番 = '" & RTrim(iH1_Hinbn) & "'"
            .DBExec
        Else
            gSL_Select = "SELECT * FROM 材料マスタ "
            gSL_Select = gSL_Select & " WHERE 材質 = '" & RTrim(iH1_Hinbn) & "'"
            If ZRMRead(gSL_Select) Then
                Do Until Not ZRM_RDSTS
                    .SQL = "DELETE FROM 材料品番マスタ "
                    .SQL = .SQL & " WHERE 材料管理番号 = " & ZRM.材料管理番号
                    .DBExec
                    
                    Call ZRMReadNext
                Loop
            End If
            Call ZRMClose
        End If
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        
        iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + CisDB.RecordCount, "###")
        
        On Error GoTo 0
    End With
    DBDelete = True
    Exit Function
DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------------+
'+      削 除 処 理            +
'+----------------------------+
Private Function DBDeleteM()
    DBDeleteM = False
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
    
    With CisDB
        On Error GoTo DBDeleteM_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 削除処理 )
        .SQL = "DELETE FROM 材料品番マスタ "
        .SQL = .SQL & " WHERE 使用品番 = '" & RTrim(VsGrid1.TextMatrix(VsGrid1.Row, SColZ(14))) & "'"
        .SQL = .SQL & "   AND 材料管理番号 = " & CisFun.Val2(VsGrid1.TextMatrix(VsGrid1.Row, SColZ(13)))
        .DBExec

        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        
        iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")
        
        On Error GoTo 0
    End With
    DBDeleteM = True
    Exit Function
DBDeleteM_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut(KB As Byte)
    DBPut = False
    Dim AddFlg  As Boolean

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    If KB = 0 Then
        With CisFun
            .MB_Lines = 4
            If SyoriKB = "ADD" Then
                .MB_MSG(2) = "     追加処理を行います      "
                .MB_Title = "追加処理"
            Else
                .MB_MSG(2) = "     修正処理を行います      "
                .MB_Title = "修正処理"
            End If
            .MB_MSG(3) = "     よろしいですか？      "
            .MB_Button = OK_CAN
            If Not .MBOX Then
                If SyoriKB = "ADD" Then
                    Call FocusIchi_Set          ' 2009.09.24 upd : iB1_Itatu.SetFocus
                Else
                    iB1_Pitch.SetFocus
                End If
                Exit Function
            End If
        End With
    End If
    On Error GoTo DBPut_Err

' マスタ仮登録
    If wsKari <> 0 Then
        If wMNTKb <> "HIN" Then
            Call ItemsClearHNM
            HNM.品番 = RTrim(iB1_Zaist)
            HNM.品名 = "*仮登録*"
            HNM.仮区分 = wsKari
            
            Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
        ' ( 追加処理 )
            HNM.作成者 = gTanto
            Call HNMInsert
        '
            CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        Else
            With CisDB
                .SQL = "番号獲得"
                .StoadoCount = 3
                .ParaNo = 0: .ParaIO = Return用
                .ParaNo = 1: .ParaIO = Input用
                .ParaNo = 2: .ParaIO = OutPut用
                .ParaNo = 3: .ParaIO = Return用
            
                .ParaNo = 1: .ParaValue = 1
            
                If Not .DBStored Then
                    With CisFun
                        .MB_Lines = 4
                        .MB_MSG(2) = "     材料管理番号 採番に失敗しました    "
                        .MB_Title = "採番処理"
                        .MB_Button = Error
                        If Not .MBOX Then GoTo DBPut_Err
                    End With
                Else
                    .ParaNo = 2: wKanri = .ParaValue
                End If
            End With
        
            Call ItemsClearZRM
            ZRM.材料管理番号 = wKanri
            ZRM.材質 = RTrim(iB1_Zaist)
            ZRM.材料名 = "*仮登録*"
            ZRM.板厚 = CisFun.Val2(iB1_Itatu)
            ZRM.幅 = CisFun.Val2(iB1_Width)
            ZRM.長さ = CisFun.Val2(iB1_Long)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT START
            ZRM.単位 = RTrim(B1lb_Tani)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT END
            If ZRM.板厚 = 0 And ZRM.幅 = 0 And ZRM.長さ = 0 Then
                ZRM.寸法区分 = 1
            End If
            ZRM.仮区分 = wsKari
            
            Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
        ' ( 追加処理 )
            ZRM.作成者 = gTanto
            Call ZRMInsert
        '
            CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        End If
    End If

'   マスタ項目移送
    If wMNTKb = "HIN" Then
        gSL_Select = "SELECT * From 材料品番マスタ"
        gSL_Select = gSL_Select & " WHERE 使用品番 = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & "   AND 材料管理番号 = " & wKanri
    Else
        gSL_Select = "SELECT * From 材料品番マスタ"
        gSL_Select = gSL_Select & " WHERE 使用品番 = '" & RTrim(iB1_Zaist) & "'"
        gSL_Select = gSL_Select & "   AND 材料管理番号 = " & wKanri
    End If
    
    If ZHMRead(gSL_Select, 1) Then
        AddFlg = False
    Else
        Call ItemsClearZHM
        AddFlg = True
        If wMNTKb = "HIN" Then
            ZHM.使用品番 = RTrim(iH1_Hinbn)
        Else
            ZHM.使用品番 = RTrim(iB1_Zaist)
        End If
        ZHM.材料管理番号 = wKanri
    End If
    
    With ZHM
        .板厚 = CisFun.Val2(iB1_Itatu)
        .幅 = CisFun.Val2(iB1_Width)
        .長さ = CisFun.Val2(iB1_Long)
        .ピッチ = CisFun.Val2(iB1_Pitch)
        .取数 = CisFun.Val2(iB1_TSur)
        .未使用量 = CisFun.Val2(iB1_MSiyou)
        .使用量 = CisFun.Val2(iB1_Siyou)
        .展開区分 = B1Chk_Tenkb.Value
        .LT = CisFun.Val2(iB1_LT)
        .有効期限開始 = RTrim(iB1_SDate)
        .有効期限終了 = RTrim(iB1_EDate)
    End With

    wBk_Kanri = wKanri
    wBk_Hinbn = ZHM.使用品番
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 追加処理 )
    If AddFlg Then
        ZHM.作成者 = gTanto
        Call ZHMInsert
        iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
    Else
' ( 更新処理 )
        ZHM.更新者 = gTanto
        Call ZHMUpdate
        iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
    End If

'
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   材料⇒品番情報   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit1()
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 UPDATE START
        'ReDim SColZ(21)
        ReDim SColZ(22)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 UPDATE END

        Call .FixedSet(中中, 右中, あり, 4, 0, "No")
        Call .FixedSet(中中, 左中, あり, 24, 1, "品番")
        Call .FixedSet(中中, 左中, あり, 6, 2, "背番号")
        Call .FixedSet(中中, 左中, あり, 24, 3, "材質")
        Call .FixedSet(中中, 中中, あり, 6, 4, "区分")
        Call .FixedSet(中中, 左中, あり, 23, 5, "サイズ")
        Call .FixedSet(中中, 右中, あり, 7, 6, "ピッチ")
        Call .FixedSet(中中, 右中, あり, 5, 7, "取数")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 UPDATE START
'        Call .FixedSet(中中, 右中, あり, 7, 8, "使用量")
'        Call .FixedSet(中中, 右中, あり, 7, 9, "屑重量")
        '================================================ *** 2008/07/15 Start
        'Call .FixedSet(中中, 右中, あり, 7, 8, "単重")
        'Call .FixedSet(中中, 右中, あり, 10, 9, "ｽｸﾗｯﾌﾟ重量")
        Call .FixedSet(中中, 右中, あり, 8, 8, "所要量")
        Call .FixedSet(中中, 右中, あり, 10, 9, "ｽｸﾗｯﾌﾟ重量")
        '================================================ *** 2008/07/15 End
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 UPDATE END
        Call .FixedSet(中中, 中中, あり, 3, 10, "展")
        Call .FixedSet(中中, 中中, あり, 2, 11, "LT")
        Call .FixedSet(中中, 左中, あり, 22, 12, "有効期限")
        Call .FixedSet(中中, 中中, あり, 0, 13, "材料管理番号")
        Call .FixedSet(中中, 中中, あり, 0, 14, "編集前品番")
        Call .FixedSet(中中, 中中, あり, 0, 15, "寸法区分")
        Call .FixedSet(中中, 中中, あり, 0, 16, "板厚")
        Call .FixedSet(中中, 中中, あり, 0, 17, "幅")
        Call .FixedSet(中中, 中中, あり, 0, 18, "長さ")
        Call .FixedSet(中中, 中中, あり, 0, 19, "有効期限開始")
        Call .FixedSet(中中, 中中, あり, 0, 20, "有効期限終了")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT START
        Call .FixedSet(中中, 中中, あり, 4, 21, "単位")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT END

'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 UPDATE START
        Call .FixedSet(中中, 左中, あり, 0, 22, "Select")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 UPDATE END
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 UPDATE START
'        If wMNTKb = "HIN" Then
'            Call .PatanSet(0, "0,3,4,5,6,7,8,9,10,12")
'        Else
'            Call .PatanSet(0, "0,1,2,5,6,7,8,9,10,12")
'        End If
        If wMNTKb = "HIN" Then
            Call .PatanSet(0, "0,3,4,5,21,6,7,8,9,10,12")
        Else
            Call .PatanSet(0, "0,1,2,5,21,6,7,8,9,10,12")
        End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 UPDATE END
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp

        SColZ(1) = .FixedGet(1)     ' 品番
        SColZ(2) = .FixedGet(2)     ' 社内背番
        SColZ(3) = .FixedGet(3)     ' 材質
        SColZ(4) = .FixedGet(4)     ' 寸法
        SColZ(5) = .FixedGet(5)     ' 板厚
        SColZ(6) = .FixedGet(6)     ' 幅
        SColZ(7) = .FixedGet(7)     ' 長さ
        SColZ(8) = .FixedGet(8)     ' ピッチ
        SColZ(9) = .FixedGet(9)     ' 取数
        SColZ(10) = .FixedGet(10)   ' 未使用
        SColZ(11) = .FixedGet(11)   ' 使用
        SColZ(12) = .FixedGet(12)   ' 有効期限
        SColZ(13) = .FixedGet(13)   ' 材料管理番号
        SColZ(14) = .FixedGet(14)   ' 編集前品番
        SColZ(15) = .FixedGet(15)   ' 寸法区分
        SColZ(16) = .FixedGet(16)   ' 有効期限開始
        SColZ(17) = .FixedGet(17)   ' 有効期限終了
        SColZ(18) = .FixedGet(18)   ' 有効期限終了
        SColZ(19) = .FixedGet(19)   ' 有効期限終了
        SColZ(20) = .FixedGet(20)   ' 有効期限終了
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT START
        SColZ(21) = .FixedGet(21)   ' 単位
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT END
        
    End With
    KCol = 13   ' 非表示項目開始ｶﾗﾑ

End Sub
'****************************************************************************************'
'*      グリッド内容セット(材料⇒品番)
'****************************************************************************************'
Private Sub GridSet_Zai()
    Dim ID          As Long
    Dim Ttl_Kensu   As Long


    gSL_Select = "SELECT HZ.*, "
    gSL_Select = gSL_Select & " ZR.材質,ZR.表示寸法,"
    gSL_Select = gSL_Select & " ZR.寸法区分,ZR.板厚,ZR.幅,ZR.長さ, "
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT START
    gSL_Select = gSL_Select & " ISNULL(SY3.値名称, '') 単位,"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT END
    gSL_Select = gSL_Select & " SY1.値名称 材種名,"
    gSL_Select = gSL_Select & " SY2.値名称 形状名, "
    gSL_Select = gSL_Select & " HN.社内背番 背番号,HN.品番編集 "
    gSL_Select = gSL_Select & " FROM 材料品番マスタ HZ"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料マスタ ZR "
    gSL_Select = gSL_Select & "   ON HZ.材料管理番号 = ZR.材料管理番号 "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "   ON HN.品番 = HZ.使用品番 "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY1 "
    gSL_Select = gSL_Select & "   ON SY1.区分名称 = '材種区分' "
    gSL_Select = gSL_Select & "  AND SY1.値 = ZR.材種 "
    gSL_Select = gSL_Select & "  AND SY1.区分タイプ = '' "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY2 "
    gSL_Select = gSL_Select & "   ON SY2.区分名称 = '材種区分' "
    gSL_Select = gSL_Select & "  AND SY2.値 = ZR.形状 "
    gSL_Select = gSL_Select & "  AND SY2.区分タイプ = '' "
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT START
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY3"
    gSL_Select = gSL_Select & "      ON SY3.区分名称 = '単位区分'"
    gSL_Select = gSL_Select & "     AND SY3.値 = ZR.単位 "
    gSL_Select = gSL_Select & "     AND SY3.区分タイプ = '' "
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT END
    If wMNTKb = "HIN" Then
        gSL_Select = gSL_Select & " WHERE HZ.使用品番 = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " ORDER BY ZR.材質,ZR.板厚,ZR.幅,ZR.長さ"
    Else
        gSL_Select = gSL_Select & " WHERE ZR.材質 = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " ORDER BY HZ.使用品番,ZR.板厚,ZR.幅,ZR.長さ"
    End If
    Call ZHMRead(gSL_Select)

' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGrid1
       .Redraw = False
        If Ttl_Kensu + 1 < CisVsGrid1.DispMax Then
            .Rows = CisVsGrid1.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0: wBK_Row = 1
    Do Until Not ZHM_RDSTS
        ID = ID + 1
            
        If wBK_Row = 1 Then
            If wBk_Kanri = ZHM.材料管理番号 And _
               wBk_Hinbn = ZHM.使用品番 Then
                wBK_Row = ID
            End If
        End If
        
        Call ZaiInfoToGrid(ID)
        
        Call ZHMReadNext
    Loop
    
    Call ZHMClose
    
    With VsGrid1
        .Row = wBK_Row: .Col = 1: .ColSel = .Cols - 1
         
        If .RowIsVisible(wBK_Row) = False Then
            .TopRow = wBK_Row
        End If
    
        .Redraw = True
    End With
End Sub
'+---------------------------------------------------------------------------------------+
'+  材料情報よりグリッドへセット
'+---------------------------------------------------------------------------------------+
Private Sub ZaiInfoToGrid(SetRow As Long)
    With VsGrid1
        For mIx = 0 To .Cols - 1
            .TextMatrix(SetRow, mIx) = ""
        Next mIx
        
        ' No
        .TextMatrix(SetRow, 0) = SetRow
        ' 品番
        .TextMatrix(SetRow, SColZ(14)) = ZHM.使用品番
        .TextMatrix(SetRow, SColZ(1)) = EditHinbn(ZHM.使用品番, ZHM.品番編集)
        ' 社内背番
        .TextMatrix(SetRow, SColZ(2)) = ZHM.背番号
        ' 材質
        .TextMatrix(SetRow, SColZ(3)) = ZHM.材質
        ' 寸法
        .TextMatrix(SetRow, SColZ(15)) = ZHM.寸法区分
        If ZHM.寸法区分 <> 0 Then
            .TextMatrix(SetRow, SColZ(4)) = "無"
        End If
        ' サイズ
        .TextMatrix(SetRow, SColZ(16)) = CisFun.RSetFld(ZHM.板厚, 7, "###0.00")
        .TextMatrix(SetRow, SColZ(17)) = CisFun.RSetFld(ZHM.幅, 7, "###0.00")
        .TextMatrix(SetRow, SColZ(18)) = CisFun.RSetFld(ZHM.長さ, 7, "###0.00")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT START
        ' 単位
        .TextMatrix(SetRow, SColZ(21)) = ZHM.単位
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT END
        ' 寸法
        .TextMatrix(SetRow, SColZ(5)) = ZHM.表示寸法
        ' ピッチ
        If ZHM.ピッチ <> 0 Then
            .TextMatrix(SetRow, SColZ(6)) = CisFun.RSetFld(ZHM.ピッチ, 7, "###0.00")
        End If
        ' 取数
        If ZHM.取数 <> 0 Then
            .TextMatrix(SetRow, SColZ(7)) = CisFun.RSetFld(ZHM.取数, 4, "####")
        End If
        ' 使用量
        If ZHM.使用量 <> 0 Then
            '================================ *** 2008/07/15 Start
            '.TextMatrix(SetRow, SColZ(8)) = CisFun.RSetFld(ZHM.使用量, 7, "###0.00")
            '.TextMatrix(SetRow, SColZ(8)) = CisFun.RSetFld(ZHM.使用量, 7, "###0.000")
            .TextMatrix(SetRow, SColZ(8)) = CisFun.RSetFld(ZHM.使用量, 7, "###0.0000")          '2010/09/17
            '================================ *** 2008/07/15 End
        End If
        ' 屑重量
        If ZHM.未使用量 <> 0 Then
            '================================ *** 2008/07/15 Start
            '.TextMatrix(SetRow, SColZ(9)) = CisFun.RSetFld(ZHM.未使用量, 7, "###0.00")
            .TextMatrix(SetRow, SColZ(9)) = CisFun.RSetFld(ZHM.未使用量, 7, "###0.000")
            '================================ *** 2008/07/15 End
        End If
        ' 展開区分
        If ZHM.展開区分 <> 0 Then
            .TextMatrix(SetRow, SColZ(10)) = "無"
        End If
        ' LT
        If ZHM.LT <> 0 Then
            .TextMatrix(SetRow, SColZ(11)) = CisFun.RSetFld(ZHM.LT, 2, "##")
        End If
        
        ' 有効期限
        .TextMatrix(SetRow, SColZ(19)) = ZHM.有効期限開始
        .TextMatrix(SetRow, SColZ(20)) = ZHM.有効期限終了
        gStr = ""
        If Trim(ZHM.有効期限開始) <> "" Then
            gStr = Mid(ZHM.有効期限開始, 1, 4) & "/" & _
                   Mid(ZHM.有効期限開始, 5, 2) & "/" & _
                   Mid(ZHM.有効期限開始, 7, 2)
        End If
        If Trim(ZHM.有効期限終了) <> "" Then
           If gStr = "" Then gStr = Space(10)
           gStr = gStr & "～"
           gStr = gStr & Mid(ZHM.有効期限終了, 1, 4) & "/" & _
                         Mid(ZHM.有効期限終了, 5, 2) & "/" & _
                         Mid(ZHM.有効期限終了, 7, 2)
        Else
           If Trim(ZHM.有効期限開始) <> "" Then
              gStr = gStr & "～" & Space(10)
           End If
        End If
        .TextMatrix(SetRow, SColZ(12)) = gStr
        
        ' 材料管理番号
        .TextMatrix(SetRow, SColZ(13)) = ZHM.材料管理番号
    End With
End Sub
'+---------------------------------------------------------------------------------------+
'+  材料情報グリッドより入力域へセット
'+---------------------------------------------------------------------------------------+
Private Sub GridToText_Zai()
    
    With VsGrid1
        If Trim(.TextMatrix(.Row, SColZ(14))) = "" Then Exit Sub
        
        If wMNTKb = "HIN" Then
            iB1_Zaist = RTrim(.TextMatrix(.Row, SColZ(3)))
            iB1_Itatu = RTrim(.TextMatrix(.Row, SColZ(4)))
            Call Get_Size(0)
        Else
            iB1_Zaist = RTrim(.TextMatrix(.Row, SColZ(14)))
            iB1_Itatu = RTrim(.TextMatrix(.Row, SColZ(2)))
            Call Get_Size(1)
        End If
        iB1_Itatu = CisFun.RSetFld(.TextMatrix(.Row, SColZ(16)), 7, "###0.00;###0.00;#")
        iB1_Width = CisFun.RSetFld(.TextMatrix(.Row, SColZ(17)), 7, "###0.00;###0.00;#")
        iB1_Long = CisFun.RSetFld(.TextMatrix(.Row, SColZ(18)), 7, "###0.00;###0.00;#")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT START
        B1lb_Tani = RTrim(.TextMatrix(.Row, SColZ(21)))
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/13 INSERT END
        
        iB1_Pitch = CisFun.RSetFld(.TextMatrix(.Row, SColZ(6)), 7, "###0.00;###0.00;#")
        iB1_TSur = CisFun.RSetFld(.TextMatrix(.Row, SColZ(7)), 4, "####")
        '======================================= *** 2008/07/15 Start
        'iB1_Siyou = CisFun.RSetFld(.TextMatrix(.Row, SColZ(8)), 7, "###0.00")
        'iB1_MSiyou = CisFun.RSetFld(.TextMatrix(.Row, SColZ(9)), 7, "###0.00")
        'iB1_Siyou = CisFun.RSetFld(.TextMatrix(.Row, SColZ(8)), 7, "###0.000")
        iB1_Siyou = CisFun.RSetFld(.TextMatrix(.Row, SColZ(8)), 7, "###0.0000")                 '2010/09/17
        iB1_MSiyou = CisFun.RSetFld(.TextMatrix(.Row, SColZ(9)), 7, "###0.000")
        '======================================= *** 2008/07/15 End

        If RTrim$(.TextMatrix(.Row, SColZ(10))) = "" Then
            B1Chk_Tenkb.Value = 0
        Else
            B1Chk_Tenkb.Value = 1
        End If

        iB1_LT = CisFun.RSetFld(.TextMatrix(.Row, SColZ(11)), 7, "###0.00")
        iB1_SDate = RTrim(.TextMatrix(.Row, SColZ(19)))
        iB1_EDate = RTrim(.TextMatrix(.Row, SColZ(20)))
    
    End With

End Sub
'+---------------------------------------+
'+  グリッド (DblClick)
'+---------------------------------------+
Private Sub vsGrid1_DblClick()

    With VsGrid1
        If ProcHB = "B2" Then

            If Trim(.TextMatrix(.Row, SColZ(14))) <> "" Then        ' 14:編集前品番
                RowColFlg = True
                Call CisVsGrid1.vsColor(消去個別)
                If SyoriKB2 = "DEL" Then
                    Call CisVsGrid1.vsColor(個別, &HC0C0FF)
                    Call ReturnPress
                Else
                    Call CisVsGrid1.vsColor(個別)
                    ProcHB = "B1"
                    Call DispChange(ProcHB)
                End If
                RowColFlg = False
            End If
        End If
    End With

End Sub
'+---------------------------------------+
'+  グリッド (KeyPress)
'+---------------------------------------+
Private Sub vsGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
       Call vsGrid1_DblClick
    End If
End Sub
'+---------------------------------------+
'+  グリッド (RowColChange)
'+---------------------------------------+
Private Sub vsGrid1_RowColChange()
    If Not FormAct Then Exit Sub
    If RTrim$(VsGrid1.Tag) <> "" Then Exit Sub
    If RowColFlg Then Exit Sub
    If VsGrid1.MouseRow = 0 And VsGrid1.Tag <> "" Then Exit Sub

    With VsGrid1
        If ProcHB = "B2" Then
            GridBack = ProcHB
            Call HeadBodyClear("B1")
             
            If Trim(.TextMatrix(.Row, SColZ(14))) <> "" Then        ' 14:編集前品番
               Call GridToText_Zai
            End If
        End If
    End With
End Sub
'******************************************'
'*****       移動選択項目選択          *****
'******************************************'
Private Sub VSGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

' Button : 1 = 左ﾎﾞﾀﾝ ,2 = 右ﾎﾞﾀﾝ
    With VsGrid1
        .Tag = ""
        If Button <> 1 Then Exit Sub
    ' 固定行以外は行わない
        If .MouseRow <> 0 Then Exit Sub
    ' 固定列は行わない
        If .MouseCol = 0 Then Exit Sub
    ' 移動元列をタグへ退避
        .Tag = Str(.MouseCol)
        .Drag 1
    End With
End Sub
'******************************************'
'*****        表 示 項 目 移 動        *****
'******************************************'
Private Sub VSGrid1_DragDrop(Source As Control, X As Single, Y As Single)
    With VsGrid1
    ' 固定列は行わない
        If .MouseCol = 0 Then Exit Sub
        If .Tag = "" Then Exit Sub
' ｸﾞﾘｯﾄﾞ表示 禁止
        .Redraw = False
' 列移動
        .ColPosition(CisFun.Val2(.Tag)) = .MouseCol

        Call CisVsGrid1.VSSort

' ｸﾞﾘｯﾄﾞ表示 解除
        CisVsGrid1.ReNumber
        .Redraw = True
            
'   < カラムワークセット >
        For gLong = 1 To VsGrid1.Cols - 1
            gInt = CisVsGrid1.FixedGetNM(gLong)
            SColZ(gInt) = gLong
        Next gLong

        .Tag = ""
    End With
End Sub

'+==========================================================+
'+      代用寸法入力グリッド処理                      +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit2()
    With CisVsGrid2
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 2, 0, "NO")
        Call .FixedSet(中中, 右中, あり, 7, 1, "板厚")
        Call .FixedSet(中中, 右中, あり, 7, 2, "幅")
        Call .FixedSet(中中, 右中, あり, 7, 3, "長さ")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 4, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3")
        .InitDisp
    End With
End Sub
Private Sub Grid_RowChange2(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
    
'テキストからグリッドに退避
    Call TextToVsGrid2(Val(iB1_DItatu.Tag))
    gInt = VsGrid2.Row
    '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange2_End
        gInt = gInt - 1
    End If
    '↓
    If KeyCode = vbKeyDown Then
        If gInt + 1 > VsGrid2.Rows - 1 Then
            If Trim(VsGrid2.TextMatrix(VsGrid2.Rows - 1, 1)) = "" Then GoTo Grid_RowChange2_End
            VsGrid2.Rows = VsGrid2.Rows + 1
        End If
        gInt = gInt + 1
    End If
    'PageUp
    If KeyCode = vbKeyPageUp Then
        If gInt - CisVsGrid2.DispMax < 1 Then
            gInt = 1
        Else
            gInt = gInt - CisVsGrid2.DispMax
        End If
    End If
    'PageDown
'    If KeyCode = vbKeyPageDown Then
'        If gInt + CisVsGrid.DispMax > Ttl_Kensu - 1 Then
'            gInt = Ttl_Kensu - 1
'        Else
'            gInt = gInt + CisVsGrid.DispMax
'        End If
'    End If
    VsGrid2.Row = gInt
'グリッドからテキストに表示
    Call VsGridToText2(VsGrid2.Row)
    
Grid_RowChange2_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid2(tRow As Integer)
    With VsGrid2
'        If Trim(iB1_ditatu) = "" Then
'            For gInt = 2 To .Cols - 1
'                .TextMatrix(tRow, gInt) = ""
'            Next gInt
'            Exit Sub
'        End If
        .TextMatrix(tRow, 0) = tRow
        .TextMatrix(tRow, 1) = iB1_DItatu
        .TextMatrix(tRow, 2) = iB1_DWidth
        .TextMatrix(tRow, 3) = iB1_DLong
    End With
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText2(tRow As Integer)
    With VsGrid2
        '代用板厚
        iB1_DItatu.Tag = tRow
        iB1_DItatu = RTrim(.TextMatrix(tRow, 1))
        iB1_DItatu.Move iB1_DItatu.Left, VsGrid2.CellTop, iB1_DItatu.Width, VsGrid2.CellHeight
        '代用幅
        iB1_DWidth = RTrim(.TextMatrix(tRow, 2))
        iB1_DWidth.Move iB1_DWidth.Left, VsGrid2.CellTop, iB1_DWidth.Width, VsGrid2.CellHeight
        '代用長さ
        iB1_DLong = RTrim(.TextMatrix(tRow, 3))
        iB1_DLong.Move iB1_DLong.Left, VsGrid2.CellTop, iB1_DLong.Width, VsGrid2.CellHeight
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet2()
    Dim ID          As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGrid2
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid2.FixedGet(CisFun.Val2(ID))
        Next ID
        If Ttl_Kensu + 1 < CisVsGrid2.DispMax Then
            .Rows = CisVsGrid2.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not ZDA_RDSTS
        With VsGrid2
            ID = ID + 1
            '№
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            '代用板厚
            .TextMatrix(ID, GridCol(1)) = Format(ZDA.板厚, "###0.00")
            '代用幅
            .TextMatrix(ID, GridCol(2)) = Format(ZDA.幅, "###0.00")
            '代用長さ
            .TextMatrix(ID, GridCol(3)) = Format(ZDA.長さ, "###0.00")
        End With
        Call ZDAReadNext(1)
    Loop
    
    With VsGrid2
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub

Private Sub vsGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown Then Exit Sub
    Call TextToVsGrid2(Val(iB1_DItatu.Tag))
    Call VsGridToText2(VsGrid2.Row)
End Sub
Private Sub vsGrid2_Click()
'    Call GridTextEnabled(True)
    Call TextToVsGrid2(Val(iB1_DItatu.Tag))
    Call VsGridToText2(VsGrid2.Row)
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub
'****************************
'*      Grid2ﾁｪｯｸ処理（代用寸法）
'****************************
Private Function Grid2Chk() As Boolean
    Dim wUke_Flg   As Boolean

    Grid2Chk = False
    
    With VsGrid2
        For gCnt1 = 1 To .Rows - 1
            If Trim(.TextMatrix(gCnt1, 1)) <> "" Or Trim(.TextMatrix(gCnt1, 2)) <> "" Or Trim(.TextMatrix(gCnt1, 3)) <> "" Then
                '寸法の重複チェック
                For gCnt2 = gCnt1 + 1 To .Rows - 1
                    If Trim(.TextMatrix(gCnt2, 1)) & Trim(.TextMatrix(gCnt2, 2)) & Trim(.TextMatrix(gCnt2, 3)) <> "" Then
                        If Trim(.TextMatrix(gCnt1, 1)) & Trim(.TextMatrix(gCnt1, 2)) & Trim(.TextMatrix(gCnt1, 3)) = Trim(.TextMatrix(gCnt2, 1)) & Trim(.TextMatrix(gCnt2, 2)) & Trim(.TextMatrix(gCnt2, 3)) Then
                            Set gObj = iB1_DItatu
                            gLong = gCnt2
                            With CisFun
                                .MB_Lines = 4
                                .MB_MSG(2) = "    " & StrConv(gCnt2, vbWide) & "行目 入力内容が重複してます。            "
                                .MB_Title = ""
                                .MB_Button = Error
                                .MBOX
                            End With
                            GoTo Grid2Chk_Err
                        End If
                    End If
                Next gCnt2
                
            End If
        Next gCnt1
    End With
    
    Grid2Chk = True
    Exit Function

Grid2Chk_Err:
    VsGrid2.Row = gLong
    Call TextToVsGrid2(Val(iB1_DItatu.Tag))
    Call VsGridToText2(VsGrid2.Row)
    gObj.SetFocus

End Function
Private Sub VsGrid2_DBInsert()
    
    With CisDB
        .SQL = "DELETE 材料代用マスタ "
        .SQL = .SQL & " WHERE 材料管理番号 = " & D_ZKanriNo
        .DBExec
    End With

    With VsGrid2
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, 1)) & Trim(.TextMatrix(gLong, 2)) & Trim(.TextMatrix(gLong, 3)) <> "" Then
               Call ItemsClearZDA
               With ZDA
                    .材料管理番号 = D_ZKanriNo
                    If wMNTKb = "HIN" Then
                        .品番 = RTrim(iH1_Hinbn)
                    Else
                        .品番 = RTrim(iB1_Zaist)
                    End If
                    .板厚 = CisFun.Val2(VsGrid2.TextMatrix(gLong, 1))
                    .幅 = CisFun.Val2(VsGrid2.TextMatrix(gLong, 2))
                    .長さ = CisFun.Val2(VsGrid2.TextMatrix(gLong, 3))
               End With
               Call ZDAInsert
            End If
        Next gLong
    End With

End Sub

'---------------------------------------------------------------' 2009.09.24 add start
'****************************************************
'*      呼出元によりフォーカス位置を決定する        *
'****************************************************
Private Sub FocusIchi_Set()
    If wMNTKb = "HIN" Or BLB_SizeBack.Enabled Then
        If SyoriKB = "ADD" Then
           iB1_Itatu.SetFocus
        Else
           iB1_Pitch.SetFocus
        End If
    Else
        iB1_Pitch.SetFocus
    End If
End Sub
'---------------------------------------------------------------' 2009.09.24 add start

