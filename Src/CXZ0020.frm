VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Begin VB.Form CXZ0020 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "在庫マスタ(材料)"
   ClientHeight    =   11190
   ClientLeft      =   75
   ClientTop       =   1530
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11190
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin VB.PictureBox PNL_Disp 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      ScaleHeight     =   435
      ScaleWidth      =   2925
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2985
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   450
         Left            =   0
         Top             =   0
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   794
         BackColor       =   12632319
         ForeColor       =   0
         Caption         =   "理論在庫数："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   90
         cPositionY      =   15
         Begin Cis3D_v60.CIS3D H1lb_Zaiko 
            Height          =   360
            Left            =   1500
            Top             =   60
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   635
            BackColor       =   12632319
            Caption         =   "###,###,##0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632319
            cBoderColor2    =   12632319
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   2
            cPositionX      =   30
         End
      End
   End
   Begin VB.PictureBox B2_Area1 
      Height          =   7125
      Left            =   1860
      ScaleHeight     =   7065
      ScaleWidth      =   11175
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3720
      Width           =   11235
      Begin vsFlexLib.vsFlexArray VsGrid1 
         Height          =   7065
         Left            =   0
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   11175
         _Version        =   196608
         _ExtentX        =   19711
         _ExtentY        =   12462
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   "№ |9999/99  |999,999,999|実棚 |XXXXXXXXX1XXXXXXXXX1XXXXXXXXX1XXXXXXXXX1|担当|9999/99/99 "
         Rows            =   21
         Cols            =   12
         BackColor       =   12648447
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   330
         SelectionMode   =   1
         BackColorAlternate=   12648447
      End
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   11640
      ScaleHeight     =   1305
      ScaleWidth      =   2625
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   690
      Width           =   2685
      Begin CisBtn_60.CisBtn PB_LOOK 
         Height          =   585
         Left            =   1740
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   60
         Top             =   660
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
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
         Left            =   900
         Top             =   660
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
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
         Left            =   1740
         Top             =   660
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
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
      Begin CisBtn_60.CisBtn PB_BACK 
         Height          =   585
         Left            =   60
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
      Begin CisBtn_60.CisBtn PB_NEXT 
         Height          =   585
         Left            =   900
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
   End
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   -30
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   420
      Width           =   855
   End
   Begin VB.PictureBox B1_Area1 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8235
      Left            =   1740
      ScaleHeight     =   8175
      ScaleWidth      =   11385
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2700
      Width           =   11445
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   8055
         Left            =   90
         Top             =   60
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   14208
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
         Begin Cis3D_v60.CIS3D BLB_Date_B 
            Height          =   885
            Left            =   4620
            Top             =   30
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   1561
            ForeColor       =   16711680
            Caption         =   "実棚日"
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
            Begin CisText_V60.CisText iB1_Day 
               Height          =   375
               Left            =   270
               TabIndex        =   7
               Top             =   390
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   661
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
            Begin Cis3D_v60.CIS3D CIS3D1 
               Height          =   285
               Left            =   780
               Top             =   480
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               ForeColor       =   0
               Caption         =   "日"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cBoderStyle     =   2
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   885
            Index           =   10
            Left            =   30
            Top             =   30
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   1561
            ForeColor       =   16711680
            Caption         =   "担 当 者"
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
            Begin CisText_V60.CisText iB1_Tanto 
               Height          =   375
               Left            =   90
               TabIndex        =   5
               Top             =   390
               Width           =   495
               _ExtentX        =   873
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
               Text            =   "XXX"
               MaxLength       =   3
            End
            Begin Cis3D_v60.CIS3D B1lb_Tanto 
               Height          =   375
               Left            =   600
               Top             =   390
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   661
               BackColor       =   16777152
               Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
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
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   885
            Index           =   1
            Left            =   2490
            Top             =   30
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   1561
            ForeColor       =   16711680
            Caption         =   "在 庫 数"
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
            Begin CisText_V60.CisText iB1_Zaiko 
               Height          =   375
               Left            =   150
               TabIndex        =   6
               Top             =   390
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   661
               cFormat         =   "##,###,##0.000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "##,###,##0.000"
               MaxLength       =   14
               cDataType       =   2
               cDataReplace    =   1
               cFcDecimal      =   3
               cFdAutoFormat   =   1
               cGFormat        =   "#######0.000"
               cILength        =   8
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   885
            Index           =   2
            Left            =   5880
            Top             =   30
            Width           =   5325
            _ExtentX        =   9393
            _ExtentY        =   1561
            ForeColor       =   16711680
            Caption         =   "変　更　事　由"
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
            Begin CisText_V60.CisText iB1_Change 
               Height          =   375
               Left            =   180
               TabIndex        =   8
               Top             =   390
               Width           =   5055
               _ExtentX        =   8916
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
               Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   40
               IMEMode         =   1
               cDataReplace    =   1
            End
         End
      End
   End
   Begin VB.PictureBox H1_Area1 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   1740
      ScaleHeight     =   1005
      ScaleWidth      =   9615
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   990
      Width           =   9675
      Begin Cis3D_v60.CIS3D UC_3D15 
         Height          =   885
         Index           =   5
         Left            =   90
         Top             =   60
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "年月度"
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
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   360
            Left            =   60
            TabIndex        =   0
            Top             =   420
            Width           =   1380
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
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   885
         Left            =   1590
         Top             =   60
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "材     質"
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
         Begin CisText_V60.CisText iH1_Zaist 
            Height          =   375
            Left            =   60
            TabIndex        =   1
            Top             =   420
            Width           =   3255
            _ExtentX        =   5741
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
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   885
         Left            =   4920
         Top             =   60
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "板 厚"
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
         Begin CisText_V60.CisText iH1_Itatu 
            Height          =   375
            Left            =   60
            TabIndex        =   2
            Top             =   420
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   661
            cFormat         =   "###0.00"
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
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   885
         Left            =   6060
         Top             =   60
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "幅"
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
         Begin CisText_V60.CisText iH1_Width 
            Height          =   375
            Left            =   60
            TabIndex        =   3
            Top             =   420
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   661
            cFormat         =   "###0.00"
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
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   885
         Left            =   7200
         Top             =   60
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "長 さ"
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
         Begin CisText_V60.CisText iH1_Long 
            Height          =   375
            Left            =   60
            TabIndex        =   4
            Top             =   420
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   661
            cFormat         =   "###0.00"
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
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   885
         Left            =   8340
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "管理単位"
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
         Begin Cis3D_v60.CIS3D H1lb_Tani 
            Height          =   375
            Left            =   150
            Top             =   420
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   661
            BackColor       =   16777152
            Caption         =   "XXXXXX"
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
   Begin VB.Timer Timer1 
      Left            =   270
      Top             =   420
   End
   Begin VB.Timer Timer2 
      Left            =   690
      Top             =   420
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
      Caption         =   "【 在庫マスタ メンテ(材料)  】"
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
         Left            =   14220
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   11670
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
Attribute VB_Name = "CXZ0020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   在庫メンテナンス（材料）
'**       フォームID    :   CXZ0020
'**       処理概要      :
'**
'**       作  成  日    :   2003/12/01  By CIS SIMPR-A
'**       変  更  日    :   2004/04/14  バグ修正
'**       変  更  日    :   2004/06/28  前進＆後退キーの処理内容が逆になっていたのを修正
'**       変  更  日    :   2004/10/06  実棚日の入力条件をなくす
'**       　  　  　                    ＆翌月在庫データが存在時、メッセージ表示（Yes or No）
'**       　  　  　                    ＆バグ修正（実棚日が当日以降でもエラーにならない！）
'**                         2010/01/15  理論在庫数表示を追加
'**       変  更  日    :   2010/07/05  変更理由入力条件をなくす
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替

    Dim CisVsGrid       As New CisVsGrid3
    
    Dim ActObj          As Object
    Dim ActRow          As Integer
    
    Dim wsDate          As String
    Dim SCol()          As Long
    
    Dim wDate           As String
    Dim wEndDay         As Byte
    
    Dim wKanri          As Long
    Dim wSize           As Byte
    Dim wTani           As String
    Dim wTaniNM         As String
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
    
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = VsGrid1
        .NoSpaceCol = 0
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 9
        .InitGet = True
        .Init
    End With
    Call GridInit
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
    wsDate = Format(Now(), "yyyymmdd")
    iH1_YM = Mid(wsDate, 1, 6)
    
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
           Case vbKeyF7:       If PB_LOOK.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
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
   If ProcHB = "B1" Or ProcHB = "B2" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      iH1_YM.SetFocus
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    
    gStr = RTrim(iH1_Zaist)
    gStr = gStr & Space(25 - Len(RTrim(iH1_Zaist)))
    gStr = gStr & Format(iH1_Itatu, "0000.00")
    gStr = gStr & Format(iH1_Width, "0000.00")
    gStr = gStr & Format(iH1_Long, "0000.00")
    
    gSL_Select = " SELECT 年月度,材質,板厚,幅,長さ"
    gSL_Select = gSL_Select & " FROM 材料在庫マスタ "
    gSL_Select = gSL_Select & " WHERE 年月度 = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "   AND 材質 + REPLICATE(' ',25 - LEN(材質)) + SIZE < '" & gStr & "'"
    gSL_Select = gSL_Select & " GROUP BY 年月度,材質,板厚,幅,長さ,SIZE "
    gSL_Select = gSL_Select & " ORDER BY 年月度,材質 Desc ,SIZE DESC"
    
    Call DBInput("Back")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Zaist.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    gStr = RTrim(iH1_Zaist)
    gStr = gStr & Space(25 - Len(RTrim(iH1_Zaist)))
    gStr = gStr & Format(iH1_Itatu, "0000.00")
    gStr = gStr & Format(iH1_Width, "0000.00")
    gStr = gStr & Format(iH1_Long, "0000.00")
    
    gSL_Select = " SELECT 年月度,材質,板厚,幅,長さ"
    gSL_Select = gSL_Select & " FROM 材料在庫マスタ "
    gSL_Select = gSL_Select & " WHERE 年月度 = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "   AND 材質 + REPLICATE(' ',25 - LEN(材質)) + SIZE > '" & gStr & "'"
    gSL_Select = gSL_Select & " GROUP BY 年月度,材質,板厚,幅,長さ,SIZE "
    gSL_Select = gSL_Select & " ORDER BY 年月度,材質,SIZE"
    Call DBInput("Next")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Zaist.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_LOOK.Tag = "iH1_Zaist" Then Call Look_ZAIST: Exit Sub
    If PB_LOOK.Tag = "iH1_Itau" Then Call Look_ZAIST: Exit Sub
    If PB_LOOK.Tag = "iH1_Width" Then Call Look_ZAIST: Exit Sub
    If PB_LOOK.Tag = "iH1_Long" Then Call Look_ZAIST: Exit Sub
    If PB_LOOK.Tag = "iB1_Tanto" Then Call Look_Tanto: Exit Sub
End Sub
'+----------------------------+
'+       品番 検索              +
'+----------------------------+
Private Sub Look_ZAIST()

    RV_Left = 0
    RV_Top = 0

    CKK0070.Show vbModal
    Unload CKK0070
    Set CKK0070 = Nothing

    If RV_Rtn Then
        iH1_Zaist = RV_Zaist
        iH1_Itatu = CisFun.RSetFld(RV_Itatu, 7, "###0.00")
        iH1_Width = CisFun.RSetFld(RV_Width, 7, "###0.00")
        iH1_Long = CisFun.RSetFld(RV_Long, 7, "###0.00")
    
        Call GetTANI
    End If

End Sub
'+----------------------------------------------------------------------------+
'+      担当者(名称マスタ) 検索
'+----------------------------------------------------------------------------+
Private Sub Look_Tanto()

    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "担当管理"
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
        iB1_Tanto = RV_Code
        Call iB1_Tanto_LostFocus
    End If
End Sub
'+-----------------------------+
'+  iH1_YM
'+-----------------------------+
Private Sub iH1_YM_GotFocus()
    PB_BACK.Visible = False
    PB_NEXT.Visible = False
    PB_LOOK.Visible = False
' *ADD* 2004/04/14
    DoEvents
End Sub
Private Sub iH1_YM_LostFocus()
' *CHG* 2004/04/14
'    PB_BACK.Visible = True
'    PB_NEXT.Visible = True
    If ProcHB = "H1" Then
        PB_BACK.Visible = True
        PB_NEXT.Visible = True
    End If
End Sub
'+-----------------------------+
'+  iH1_Zaist
'+-----------------------------+
Private Sub iH1_Zaist_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH1_Zaist_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If

    Call GetTANI
End Sub
'+-----------------------------+
'+  iH1_Itatu
'+-----------------------------+
Private Sub iH1_Itatu_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH1_Itatu_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If

    Call GetTANI
End Sub
'+-----------------------------+
'+  iH1_Width
'+-----------------------------+
Private Sub iH1_Width_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH1_Width_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If

    Call GetTANI
End Sub
'+-----------------------------+
'+  iH1_Long
'+-----------------------------+
Private Sub iH1_Long_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH1_Long_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If

    Call GetTANI
End Sub
'+-----------------------------+
'+  iB1_Tanto
'+-----------------------------+
Private Sub iB1_Tanto_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iB1_Tanto_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If

    B1lb_Tanto = ""
    If MeisyoGet("担当管理", RTrim(iB1_Tanto)) Then
        B1lb_Tanto = RTrim(SYM_Meisyo)
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
    H1Mode = False
    H1Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
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
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    
    PB_LOOK.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    
'*--------------------------------------------------*<< 2004/10/06 >> Delete Start
' 実棚日
'   If Pro = "B1" Then
'       If RTrim(iH1_YM) <> Mid(wsDate, 1, 6) Then
'           BLB_Date_B.Enabled = False
'       Else
'           BLB_Date_B.Enabled = True
'       End If
'   End If
'*--------------------------------------------------*<< 2004/10/06 >> Delete End
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
'                .Enabled = True
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
        CisVsGrid.Clear
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
    
' 年月度
    If RTrim(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    年月度を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_YM.SetFocus
       Exit Function
    End If
' 品番
    If RTrim(iH1_Zaist) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    材質を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Zaist.SetFocus
       Exit Function
    Else
        gSL_Select = "SELECT * FROM 材料マスタ"
        gSL_Select = gSL_Select & " WHERE 材質 = '" & RTrim(iH1_Zaist) & "'"
        gSL_Select = gSL_Select & "   AND 板厚 = " & CisFun.Val2(iH1_Itatu)
        gSL_Select = gSL_Select & "   AND 幅 = " & CisFun.Val2(iH1_Width)
        gSL_Select = gSL_Select & "   AND 長さ = " & CisFun.Val2(iH1_Long)
        If Not ZRMRead(gSL_Select, 1) Then
            With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    材料マスタ  未登録            "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
            End With
            iH1_Zaist.SetFocus
            Exit Function
        End If
    End If
'*--------------------------------------------------*<< 2004/10/06 >> Insert Start
'   >> 翌月在庫データの存在ﾁｪｯｸ
    gSL_Select = "select * from 材料在庫マスタ"
    gSL_Select = gSL_Select & " where 年月度 > '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "   and 材料管理番号 = " & ZRM.材料管理番号 & ""
    If ZZMRead(gSL_Select, 1) Then
        With CisFun
           .MB_Lines = 4
           .MB_MSG(1) = "    " & CisFun.Mid2(ZZM.年月度, 1, 4) & "年" & CisFun.Mid2(ZZM.年月度, 5, 2) & "月度のデータが在庫マスタに存在します。            "
           .MB_MSG(3) = "          処理を続行しますか？            "
           .MB_Title = "処理確認"
           .MB_Button = Yes_No
           If Not .MBOX Then Exit Function
        End With
    End If
'*--------------------------------------------------*<< 2004/10/06 >> Insert End

    wKanri = ZRM.材料管理番号
    wSize = ZRM.寸法区分
    If ZRM.在庫区分 = 0 Then
        wTani = ZRM.単位
    Else
        wTani = ZRM.手配単位
    End If
    wTaniNM = ""
    If MeisyoGet("単位区分", wTani) Then: H1lb_Tani = SYM_Meisyo
' BODY部 セット
    Call BodySet("")
    
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
' 担当者
    If RTrim(iB1_Tanto) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    担当者を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_Tanto.SetFocus
       Exit Function
    Else
        If Not MeisyoGet("担当管理", RTrim(iB1_Tanto)) Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    担 当 者  未 登 録                "
                .MB_MSG(4) = "     【 名称マスタ 】                "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB1_Tanto.SetFocus
            Exit Function
        End If
    End If

' 実棚日
    wDate = RTrim(iH1_YM) & "01"
    With CisFun
        .Date_RsFormat = fDD
        .Date_End wDate
        wEndDay = .Val2(.Date_Result)
    End With
    If CisFun.Val2(iB1_Day) > wEndDay Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    正しい実棚日を入力して下さい        "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_Day.SetFocus
        Exit Function
    End If
'*--------------------------------------------------*<< 2004/10/06 >> Update Start
'   If CisFun.Val2(iB1_Day) > Format(Now(), "dd") Then
'       With CisFun
'           .MB_Lines = 4
'           .MB_MSG(2) = "    正しい実棚日を入力して下さい        "
'           .MB_Title = ""
'           .MB_Button = Error
'           .MBOX
'       End With
'       iB1_Day.SetFocus
'       Exit Function
'   End If
    If (iH1_YM & Format(CisFun.Val2(iB1_Day), "@@")) > Format(Now(), "yyyymmdd") Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    実棚日は今日迄の日で入力して下さい        "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_Day.SetFocus
        Exit Function
    End If
'*--------------------------------------------------*<< 2004/10/06 >> Update Start

' 変更事由
'======================== 2010/07/05
'    If RTrim(iB1_Change) = "" Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    変更事由を入力して下さい        "
'            .MB_Title = ""
'            .MB_Button = Error
'            .MBOX
'        End With
'        iB1_Change.SetFocus
'        Exit Function
'    End If
            
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
    
    If Not ZZMRead(gSL_Select, 1) Then
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
        iH1_YM = RTrim(ZZM.年月度)
        iH1_Zaist = RTrim(ZZM.材質)
        iH1_Itatu = CisFun.RSetFld(ZZM.板厚, 7, "###0.00")
        iH1_Width = CisFun.RSetFld(ZZM.幅, 7, "###0.00")
        iH1_Long = CisFun.RSetFld(ZZM.長さ, 7, "###0.00")
        Call GetTANI
        Call BodySet("")
''''Call GetZaikoSu             ' 2010.01.15 add
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
    Static wsBDFlg  As Boolean

'
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add RTrim(iH1_YM), M, -12, ResultDate
        gStr = .Date_Result
    End With

' 在庫履歴マスタ 情報
    gSL_Select = "Select * from 材料在庫履歴マスタ "
    gSL_Select = gSL_Select & " Where 材料管理番号 = " & wKanri
    gSL_Select = gSL_Select & "   AND 年月度 BETWEEN '" & RTrim(gStr) & "' AND '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & " ORDER BY 年月度 DESC,入力日付 DESC "
    If ZZRRead(gSL_Select) Then
       Call GridSet
    End If
Call GetZaikoSu             ' 2010.01.15 add
    
    Call ZZMClose
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
    DBPut = False
    Dim AddFlg  As Boolean
    Dim wZAIKO  As Long
    Dim wJDate      As String

    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     更新処理を行います      "
        .MB_Title = "更新処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err

    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

'   在庫マスタ 削除
    gSL_Select = "DELETE 材料在庫マスタ "
    gSL_Select = gSL_Select & " WHERE 年月度 = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "   And 材料管理番号 = " & wKanri
    With CisDB
        .ConnectNo = 1
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    
' 在庫マスタ 作成
    Call ItemsClearZZM
    With ZZM
        .年月度 = RTrim(iH1_YM)
        .材料管理番号 = wKanri
        .寸法区分 = wSize
        .材質 = RTrim(iH1_Zaist)
        .板厚 = CisFun.Val2(iH1_Itatu)
        .幅 = CisFun.Val2(iH1_Width)
        .長さ = CisFun.Val2(iH1_Long)
        .在庫数 = CisFun.Val2(iB1_Zaiko)
        wJDate = ""
        If CisFun.Val2(iB1_Day) <> 0 Then
            wJDate = RTrim(iH1_YM) & Format(iB1_Day, "00")
        End If
        .実棚日 = RTrim(wJDate)
        .変更事由 = RTrim(iB1_Change)
        .作成区分 = 1
    End With
    ZZM.作成者 = gTanto
    Call ZZMInsert

' 履歴区分更新
    gSL_Select = " UPDATE 材料在庫履歴マスタ SET "
    gSL_Select = gSL_Select & " 履歴区分 = 1 "
    gSL_Select = gSL_Select & " WHERE 年月度 = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "   AND 材料管理番号 = " & wKanri
    gSL_Select = gSL_Select & "   AND 実棚日 = '" & RTrim(wJDate) & "'"
    With CisDB
        .ConnectNo = 1
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With


' 在庫履歴マスタ 作成
    Call ItemsClearZZR
    With ZZR
        .年月度 = RTrim(iH1_YM)
        .材料管理番号 = wKanri
        .寸法区分 = wSize
        .材質 = RTrim(iH1_Zaist)
        .板厚 = CisFun.Val2(iH1_Itatu)
        .幅 = CisFun.Val2(iH1_Width)
        .長さ = CisFun.Val2(iH1_Long)
        .入力日付 = Now()
        .入力担当者 = RTrim(iB1_Tanto)
        .在庫数 = CisFun.Val2(iB1_Zaiko)
        If CisFun.Val2(iB1_Day) <> 0 Then
            .実棚日 = RTrim(iH1_YM) & Format(iB1_Day, "00")
        End If
        .変更事由 = RTrim(iB1_Change)
        .作成区分 = 1
    End With
    ZZR.作成者 = gTanto
    Call ZZRInsert
    
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
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    ReDim SCol(10)
    
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "NO")
        Call .FixedSet(中中, 中中, あり, 4, 1, "区分")
        Call .FixedSet(中中, 中中, あり, 9, 2, "年月度")
        Call .FixedSet(中中, 右中, あり, 14, 3, "在庫数")
        Call .FixedSet(中中, 中中, あり, 5, 4, "実棚")
        Call .FixedSet(中中, 左中, あり, 37, 5, "変更事由")
        Call .FixedSet(中中, 左中, あり, 5, 6, "担当")
        Call .FixedSet(中中, 中中, あり, 12, 7, "変更日")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 8, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,2,3,4,5,6,7")
        .InitDisp
    
        SCol(1) = .FixedGet(1)  ' 区分
        SCol(2) = .FixedGet(2)  ' 年月度
        SCol(3) = .FixedGet(3)  ' 在庫数
        SCol(4) = .FixedGet(4)  ' 実棚日
        SCol(5) = .FixedGet(5)  ' 変更事由
        SCol(6) = .FixedGet(6)  ' 担当
        SCol(7) = .FixedGet(7)  ' 変更日
    End With
'    KCol = 8   ' 非表示項目開始ｶﾗﾑ
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGrid1
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(CisFun.Val2(ID))
        Next ID
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    
    ID = 0
    Do Until Not ZZR_RDSTS
        With VsGrid1
            ID = ID + 1
            ' №
            .TextMatrix(ID, 0) = Format(ID, "###")
            ' 年月度
            .TextMatrix(ID, SCol(2)) = Mid(ZZR.年月度, 1, 4) & "/" & Mid(ZZR.年月度, 5, 2)
            ' 在庫数
            .TextMatrix(ID, SCol(3)) = CisFun.RSetFld(ZZR.在庫数, 14, "##,###,##0.000")
            ' 実棚日
            If RTrim(ZZR.実棚日) <> "" Then: .TextMatrix(ID, SCol(4)) = Mid(ZZR.実棚日, 7, 2)
            ' 変更事由
            .TextMatrix(ID, SCol(5)) = RTrim(ZZR.変更事由)
            ' 担当者
            .TextMatrix(ID, SCol(6)) = RTrim(ZZR.入力担当者)
            ' 変更日
            .TextMatrix(ID, SCol(7)) = Format(ZZR.入力日付, "yyyy/mm/dd")
        
        End With
        Call ZZRReadNext
    Loop
    
    With VsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub

'*******************************************************************
'*      管理単位獲得
'*******************************************************************
Private Sub GetTANI()

    gSL_Select = "SELECT * FROM 材料マスタ "
    gSL_Select = gSL_Select & " WHERE 材質 = '" & RTrim(iH1_Zaist) & "'"
    gSL_Select = gSL_Select & "   AND 板厚 = " & CisFun.Val2(iH1_Itatu)
    gSL_Select = gSL_Select & "   AND 幅 = " & CisFun.Val2(iH1_Width)
    gSL_Select = gSL_Select & "   AND 長さ = " & CisFun.Val2(iH1_Long)
    If Not ZRMRead(gSL_Select, 1) Then Exit Sub
    
    wKanri = ZRM.材料管理番号
    wSize = ZRM.寸法区分
    If ZRM.在庫区分 = 0 Then
        wTani = ZRM.単位
    Else
        wTani = ZRM.手配単位
    End If
    H1lb_Tani = ""
    If MeisyoGet("単位区分", wTani) Then: H1lb_Tani = SYM_Meisyo
        
End Sub

'-----------------------------------------------------------------------------------------------------------' 2010.01.15 add start
Private Function GetZaikoSu() As Boolean
    GetZaikoSu = False
    Dim wKanriNo    As Integer: wKanriNo = 0
    Dim wZaikoSu    As Long:    wZaikoSu = 0
    Dim wMasterSu   As Long:    wMasterSu = 0

    H1lb_Zaiko.Caption = ""
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT * FROM 材料在庫マスタ"
    gSL_Select = gSL_Select & " WHERE 年月度 = '" & Trim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "   AND 材質   = '" & Trim(iH1_Zaist) & "'"
    gSL_Select = gSL_Select & "   AND 板厚   =  " & Trim(Str(CisFun.Val2(iH1_Itatu)))
    gSL_Select = gSL_Select & "   AND 幅     =  " & Trim(Str(CisFun.Val2(iH1_Width)))
    gSL_Select = gSL_Select & "   AND 長さ   =  " & Trim(Str(CisFun.Val2(iH1_Long)))
    
    
    If Not ZZMRead(gSL_Select) Then
        wMasterSu = 0
        '---在庫マスタ：無
''        gSL_Select = ""
''        gSL_Select = gSL_Select & "SELECT ISNULL(SUM(SN.入荷数),0) + ISNULL(SUM(JK1.実績数),0) - ISNULL(SUM(ZS1.使用量),0) - ISNULL(SUM(JK2.実績数),0) AS 入荷数"
''        gSL_Select = gSL_Select & "      ,ISNULL(SUM(SN.入荷数),0) AS 入荷数入, ISNULL(SUM(JK1.実績数),0) AS 実績数入, ISNULL(SUM(ZS1.使用量),0) AS 使用量出, ISNULL(SUM(JK2.実績数),0) AS 実績数出"
''        gSL_Select = gSL_Select & " FROM 材料入荷テーブル SN"
''        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料実績管理テーブル JK1"
''        gSL_Select = gSL_Select & "        ON  SUBSTRING(JK1.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
''        gSL_Select = gSL_Select & "        AND JK1.材質     = SN.材質"
''        gSL_Select = gSL_Select & "        AND JK1.板厚     = SN.板厚"
''        gSL_Select = gSL_Select & "        AND JK1.幅       = SN.幅"
''        gSL_Select = gSL_Select & "        AND JK1.長さ     = SN.長さ"
''        gSL_Select = gSL_Select & "        AND JK1.実績区分 = 1"
''        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料実績管理テーブル JK2"
''        gSL_Select = gSL_Select & "        ON  SUBSTRING(JK2.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
''        gSL_Select = gSL_Select & "        AND JK2.材質     = SN.材質"
''        gSL_Select = gSL_Select & "        AND JK2.板厚     = SN.板厚"
''        gSL_Select = gSL_Select & "        AND JK2.幅       = SN.幅"
''        gSL_Select = gSL_Select & "        AND JK2.長さ     = SN.長さ"
''        gSL_Select = gSL_Select & "        AND JK2.実績区分 <> 1"
''        gSL_Select = gSL_Select & "        AND JK2.入力種類 <> 'A'"
''        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料使用量テーブル ZS1"
''        gSL_Select = gSL_Select & "        ON  SUBSTRING(ZS1.作業日,1,6) = SUBSTRING(SN.入荷日,1,6)"
''        gSL_Select = gSL_Select & "        AND ZS1.材質     = SN.材質"
''        gSL_Select = gSL_Select & "        AND ZS1.板厚     = SN.板厚"
''        gSL_Select = gSL_Select & "        AND ZS1.幅       = SN.幅"
''        gSL_Select = gSL_Select & "        AND ZS1.長さ     = SN.長さ"
'''        gSL_Select = gSL_Select & "        AND ZS1.実績区分 = 1"
''        gSL_Select = gSL_Select & " WHERE SUBSTRING(SN.入荷日,1,6) >= '" & Trim(iH1_YM) & "'"
''        gSL_Select = gSL_Select & "   AND SN.材質 = '" & Trim(iH1_Zaist) & "'"
''        gSL_Select = gSL_Select & "   AND SN.板厚 =  " & Trim(Str(CisFun.Val2(iH1_Itatu)))
''        gSL_Select = gSL_Select & "   AND SN.幅   =  " & Trim(Str(CisFun.Val2(iH1_Width)))
''        gSL_Select = gSL_Select & "   AND SN.長さ =  " & Trim(Str(CisFun.Val2(iH1_Long)))
        gSL_Select = ""
        gSL_Select = gSL_Select & ""
        gSL_Select = gSL_Select & "SELECT SUBSTRING(SN.入荷日,1,6) AS 入荷日, SN.材質, SN.板厚, SN.幅, SN.長さ"
        gSL_Select = gSL_Select & "      ,ISNULL(SN1.入荷数,0) + ISNULL(JK1.実績数,0) - ISNULL(ZS1.使用量,0) - ISNULL(JK2.実績数,0) AS 入荷数"
        gSL_Select = gSL_Select & "      ,ISNULL(SN1.入荷数,0) AS 入荷数入, ISNULL(JK1.実績数,0) AS 実績数入"
        gSL_Select = gSL_Select & "      ,ISNULL(JK2.実績数,0) AS 実績数出, ISNULL(ZS1.使用量,0) AS 使用量出"
        gSL_Select = gSL_Select & " FROM 材料入荷テーブル SN"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(S1.入荷日,1,6) AS 入荷日, S1.材質, S1.板厚, S1.幅, S1.長さ, SUM(S1.入荷数) AS 入荷数"
        gSL_Select = gSL_Select & "                         FROM  材料入荷テーブル S1"
        gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(S1.入荷日,1,6), S1.材質, S1.板厚, S1.幅, S1.長さ) SN1"
        gSL_Select = gSL_Select & "        ON  SUBSTRING(SN1.入荷日,1,6) = SUBSTRING(SN.入荷日,1,6)"
        gSL_Select = gSL_Select & "        AND SN1.材質     = SN.材質"
        gSL_Select = gSL_Select & "        AND SN1.板厚     = SN.板厚"
        gSL_Select = gSL_Select & "        AND SN1.幅       = SN.幅"
        gSL_Select = gSL_Select & "        AND SN1.長さ     = SN.長さ"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(J1.実績日,1,6) AS 実績日, J1.材質, J1.板厚, J1.幅, J1.長さ, SUM(J1.実績数) AS 実績数"
        gSL_Select = gSL_Select & "                         FROM  材料実績管理テーブル J1"
        gSL_Select = gSL_Select & "                         WHERE J1.実績区分 = 1"
        gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J1.実績日,1,6), J1.材質, J1.板厚, J1.幅, J1.長さ) JK1"
        gSL_Select = gSL_Select & "        ON  SUBSTRING(JK1.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
        gSL_Select = gSL_Select & "        AND JK1.材質     = SN.材質"
        gSL_Select = gSL_Select & "        AND JK1.板厚     = SN.板厚"
        gSL_Select = gSL_Select & "        AND JK1.幅       = SN.幅"
        gSL_Select = gSL_Select & "        AND JK1.長さ     = SN.長さ"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(J2.実績日,1,6) AS 実績日, J2.材質, J2.板厚, J2.幅, J2.長さ, SUM(J2.実績数) AS 実績数"
        gSL_Select = gSL_Select & "                         FROM  材料実績管理テーブル J2"
        gSL_Select = gSL_Select & "                         WHERE J2.実績区分 <> 1 AND J2.入力種類 <> 'A'"
        gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J2.実績日,1,6), J2.材質, J2.板厚, J2.幅, J2.長さ) JK2"
        gSL_Select = gSL_Select & "        ON  SUBSTRING(JK2.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
        gSL_Select = gSL_Select & "        AND JK2.材質     = SN.材質"
        gSL_Select = gSL_Select & "        AND JK2.板厚     = SN.板厚"
        gSL_Select = gSL_Select & "        AND JK2.幅       = SN.幅"
        gSL_Select = gSL_Select & "        AND JK2.長さ     = SN.長さ"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(Z1.作業日,1,6) AS 作業日, Z1.材質, Z1.板厚, Z1.幅, Z1.長さ, SUM(Z1.使用量) AS 使用量"
        gSL_Select = gSL_Select & "                         FROM  材料使用量テーブル Z1"
        gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(z1.作業日,1,6), Z1.材質, Z1.板厚, Z1.幅, Z1.長さ) ZS1"
        gSL_Select = gSL_Select & "        ON  SUBSTRING(ZS1.作業日,1,6) = SUBSTRING(SN.入荷日,1,6)"
        gSL_Select = gSL_Select & "        AND ZS1.材質     = SN.材質"
        gSL_Select = gSL_Select & "        AND ZS1.板厚     = SN.板厚"
        gSL_Select = gSL_Select & "        AND ZS1.幅       = SN.幅"
        gSL_Select = gSL_Select & "        AND ZS1.長さ     = SN.長さ"
        gSL_Select = gSL_Select & " WHERE SUBSTRING(SN.入荷日,1,6) >= '" & Trim(iH1_YM) & "'"
        gSL_Select = gSL_Select & "   AND SN.材質 = '" & Trim(iH1_Zaist) & "'"
''        gSL_Select = gSL_Select & "   AND SN.板厚 =  " & Trim(iH1_Itatu) & " "
''        gSL_Select = gSL_Select & "   AND SN.幅   =  " & Trim(iH1_Width) & " "
''        gSL_Select = gSL_Select & "   AND SN.長さ =  " & Trim(iH1_Long) & " "
        gSL_Select = gSL_Select & "   AND SN.板厚 =  " & CisFun.Val2(iH1_Itatu) & " "
        gSL_Select = gSL_Select & "   AND SN.幅   =  " & CisFun.Val2(iH1_Width) & " "
        gSL_Select = gSL_Select & "   AND SN.長さ =  " & CisFun.Val2(iH1_Long) & " "
        gSL_Select = gSL_Select & " GROUP BY SUBSTRING(SN.入荷日,1,6), SN.材質,   SN.板厚,   SN.幅, SN.長さ"
        gSL_Select = gSL_Select & "         ,SN1.入荷数, JK1.実績数  , JK2.実績数,ZS1.使用量"
        gSL_Select = gSL_Select & " ORDER BY SUBSTRING(SN.入荷日,1,6)"
    Else
        '---在庫マスタ：有
        wKanriNo = ZZM.材料管理番号
        wMasterSu = ZZM.在庫数
        If Trim(ZZM.実棚日) = "" Then
            '---実棚日：空白
''            gSL_Select = ""
''            gSL_Select = gSL_Select & "SELECT ISNULL(SUM(SN.入荷数),0) + ISNULL(SUM(JK1.実績数),0) - ISNULL(SUM(JK2.実績数),0) AS 入荷数"
''            gSL_Select = gSL_Select & "      ,ISNULL(SUM(SN.入荷数),0) AS 入荷数入, ISNULL(SUM(JK1.実績数),0) AS 実績数入, ISNULL(SUM(JK2.実績数),0) AS 実績数出"
''            gSL_Select = gSL_Select & " FROM 材料入荷テーブル SN"
''            gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料実績管理テーブル JK1"
''            gSL_Select = gSL_Select & "        ON  SUBSTRING(JK1.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
''            gSL_Select = gSL_Select & "        AND JK1.材質     = SN.材質"
''            gSL_Select = gSL_Select & "        AND JK1.板厚     = SN.板厚"
''            gSL_Select = gSL_Select & "        AND JK1.幅       = SN.幅"
''            gSL_Select = gSL_Select & "        AND JK1.長さ     = SN.長さ"
''            gSL_Select = gSL_Select & "        AND JK1.実績区分 = 1"
''            gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料実績管理テーブル JK2"
''            gSL_Select = gSL_Select & "        ON  SUBSTRING(JK2.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
''            gSL_Select = gSL_Select & "        AND JK2.材質     = SN.材質"
''            gSL_Select = gSL_Select & "        AND JK2.板厚     = SN.板厚"
''            gSL_Select = gSL_Select & "        AND JK2.幅       = SN.幅"
''            gSL_Select = gSL_Select & "        AND JK2.長さ     = SN.長さ"
''            gSL_Select = gSL_Select & "        AND JK2.実績区分 <> 1"
''            gSL_Select = gSL_Select & "        AND JK2.入力種類 <> 'A'"
''            gSL_Select = gSL_Select & " WHERE SUBSTRING(SN.入荷日,1,6) >= '" & Trim(ZZM.年月度) & "'"
''            gSL_Select = gSL_Select & "   AND SN.材料管理番号 = " & Trim(Str(wKanriNo))
''''''        gSL_Select = gSL_Select & "   AND SN.材質 = '" & Trim(iH1_Zaist) & "'"
''''''        gSL_Select = gSL_Select & "   AND SN.板厚 =  " & Trim(Str(CisFun.Val2(iH1_Itatu)))
''''''        gSL_Select = gSL_Select & "   AND SN.幅   =  " & Trim(Str(CisFun.Val2(iH1_Width)))
''''''        gSL_Select = gSL_Select & "   AND SN.長さ =  " & Trim(Str(CisFun.Val2(iH1_Long)))
            gSL_Select = ""
            gSL_Select = gSL_Select & ""
            gSL_Select = gSL_Select & "SELECT SUBSTRING(SN.入荷日,1,6) AS 入荷日, SN.材質, SN.板厚, SN.幅, SN.長さ"
            gSL_Select = gSL_Select & "      ,ISNULL(SN1.入荷数,0) + ISNULL(JK1.実績数,0) - ISNULL(ZS1.使用量,0) - ISNULL(JK2.実績数,0) AS 入荷数"
            gSL_Select = gSL_Select & "      ,ISNULL(SN1.入荷数,0) AS 入荷数入, ISNULL(JK1.実績数,0) AS 実績数入"
            gSL_Select = gSL_Select & "      ,ISNULL(JK2.実績数,0) AS 実績数出, ISNULL(ZS1.使用量,0) AS 使用量出"
            gSL_Select = gSL_Select & " FROM 材料入荷テーブル SN"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(S1.入荷日,1,6) AS 入荷日, S1.材質, S1.板厚, S1.幅, S1.長さ, SUM(S1.入荷数) AS 入荷数"
            gSL_Select = gSL_Select & "                         FROM  材料入荷テーブル S1"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(S1.入荷日,1,6), S1.材質, S1.板厚, S1.幅, S1.長さ) SN1"
            gSL_Select = gSL_Select & "        ON  SUBSTRING(SN1.入荷日,1,6) = SUBSTRING(SN.入荷日,1,6)"
            gSL_Select = gSL_Select & "        AND SN1.材質     = SN.材質"
            gSL_Select = gSL_Select & "        AND SN1.板厚     = SN.板厚"
            gSL_Select = gSL_Select & "        AND SN1.幅       = SN.幅"
            gSL_Select = gSL_Select & "        AND SN1.長さ     = SN.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(J1.実績日,1,6) AS 実績日, J1.材質, J1.板厚, J1.幅, J1.長さ, SUM(J1.実績数) AS 実績数"
            gSL_Select = gSL_Select & "                         FROM  材料実績管理テーブル J1"
            gSL_Select = gSL_Select & "                         WHERE J1.実績区分 = 1"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J1.実績日,1,6), J1.材質, J1.板厚, J1.幅, J1.長さ) JK1"
            gSL_Select = gSL_Select & "        ON  SUBSTRING(JK1.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
            gSL_Select = gSL_Select & "        AND JK1.材質     = SN.材質"
            gSL_Select = gSL_Select & "        AND JK1.板厚     = SN.板厚"
            gSL_Select = gSL_Select & "        AND JK1.幅       = SN.幅"
            gSL_Select = gSL_Select & "        AND JK1.長さ     = SN.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(J2.実績日,1,6) AS 実績日, J2.材質, J2.板厚, J2.幅, J2.長さ, SUM(J2.実績数) AS 実績数"
            gSL_Select = gSL_Select & "                         FROM  材料実績管理テーブル J2"
            gSL_Select = gSL_Select & "                         WHERE J2.実績区分 <> 1 AND J2.入力種類 <> 'A'"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J2.実績日,1,6), J2.材質, J2.板厚, J2.幅, J2.長さ) JK2"
            gSL_Select = gSL_Select & "        ON  SUBSTRING(JK2.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
            gSL_Select = gSL_Select & "        AND JK2.材質     = SN.材質"
            gSL_Select = gSL_Select & "        AND JK2.板厚     = SN.板厚"
            gSL_Select = gSL_Select & "        AND JK2.幅       = SN.幅"
            gSL_Select = gSL_Select & "        AND JK2.長さ     = SN.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(Z1.作業日,1,6) AS 作業日, Z1.材質, Z1.板厚, Z1.幅, Z1.長さ, SUM(Z1.使用量) AS 使用量"
            gSL_Select = gSL_Select & "                         FROM  材料使用量テーブル Z1"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(z1.作業日,1,6), Z1.材質, Z1.板厚, Z1.幅, Z1.長さ) ZS1"
            gSL_Select = gSL_Select & "        ON  SUBSTRING(ZS1.作業日,1,6) = SUBSTRING(SN.入荷日,1,6)"
            gSL_Select = gSL_Select & "        AND ZS1.材質     = SN.材質"
            gSL_Select = gSL_Select & "        AND ZS1.板厚     = SN.板厚"
            gSL_Select = gSL_Select & "        AND ZS1.幅       = SN.幅"
            gSL_Select = gSL_Select & "        AND ZS1.長さ     = SN.長さ"
            gSL_Select = gSL_Select & " WHERE SUBSTRING(SN.入荷日,1,6) >= '" & Trim(iH1_YM) & "'"
            gSL_Select = gSL_Select & "   AND SN.材料管理番号 = " & Trim(Str(wKanriNo))
''''        gSL_Select = gSL_Select & "   AND SN.材質 = '" & Trim(iH1_Zaist) & "'"
''''        gSL_Select = gSL_Select & "   AND SN.板厚 =  " & Trim(iH1_Itatu) & " "
''''        gSL_Select = gSL_Select & "   AND SN.幅   =  " & Trim(iH1_Width) & " "
''''        gSL_Select = gSL_Select & "   AND SN.長さ =  " & Trim(iH1_Long) & " "
            gSL_Select = gSL_Select & " GROUP BY SUBSTRING(SN.入荷日,1,6), SN.材質,   SN.板厚,   SN.幅, SN.長さ"
            gSL_Select = gSL_Select & "         ,SN1.入荷数, JK1.実績数  , JK2.実績数,ZS1.使用量"
            gSL_Select = gSL_Select & " ORDER BY SUBSTRING(SN.入荷日,1,6)"
        Else
            '---実棚日：入力済
''            gSL_Select = ""
''            gSL_Select = gSL_Select & "SELECT ISNULL(SUM(SN.入荷数),0) + ISNULL(SUM(JK1.実績数),0) - ISNULL(SUM(JK2.実績数),0) AS 入荷数"
''            gSL_Select = gSL_Select & "      ,ISNULL(SUM(SN.入荷数),0) AS 加工数入, ISNULL(SUM(JK1.実績数),0) AS 実績数入, ISNULL(SUM(JK2.実績数),0) AS 実績数出"
''            gSL_Select = gSL_Select & " FROM 材料入荷テーブル SN"
''            gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料実績管理テーブル JK1"
''            gSL_Select = gSL_Select & "        ON  JK1.実績日   > '" & Trim(ZZM.実棚日) & "'"
''            gSL_Select = gSL_Select & "        AND JK1.材質     = SN.材質"
''            gSL_Select = gSL_Select & "        AND JK1.板厚     = SN.板厚"
''            gSL_Select = gSL_Select & "        AND JK1.幅       = SN.幅"
''            gSL_Select = gSL_Select & "        AND JK1.長さ     = SN.長さ"
''            gSL_Select = gSL_Select & "        AND JK1.実績区分 = 1"
''            gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料実績管理テーブル JK2"
''            gSL_Select = gSL_Select & "        ON  JK2.実績日   > '" & Trim(ZZM.実棚日) & "'"
''            gSL_Select = gSL_Select & "        AND JK2.材質     = SN.材質"
''            gSL_Select = gSL_Select & "        AND JK2.板厚     = SN.板厚"
''            gSL_Select = gSL_Select & "        AND JK2.幅       = SN.幅"
''            gSL_Select = gSL_Select & "        AND JK2.長さ     = SN.長さ"
''            gSL_Select = gSL_Select & "        AND JK2.実績区分 <> 1"
''            gSL_Select = gSL_Select & "        AND JK2.入力種類 <> 'A'"
''            gSL_Select = gSL_Select & "  WHERE SN.入荷日 > '" & Trim(ZZM.実棚日) & "'"
''            gSL_Select = gSL_Select & "   AND SN.材料管理番号 = " & Trim(Str(wKanriNo))
''''''        gSL_Select = gSL_Select & "   AND SN.材質 = '" & Trim(iH1_Zaist) & "'"
''''''        gSL_Select = gSL_Select & "   AND SN.板厚 =  " & Trim(Str(CisFun.Val2(iH1_Itatu)))
''''''        gSL_Select = gSL_Select & "   AND SN.幅   =  " & Trim(Str(CisFun.Val2(iH1_Width)))
''''''        gSL_Select = gSL_Select & "   AND SN.長さ =  " & Trim(Str(CisFun.Val2(iH1_Long)))
            gSL_Select = ""
            gSL_Select = gSL_Select & ""
            gSL_Select = gSL_Select & "SELECT SUBSTRING(SN.入荷日,1,6) AS 入荷日, SN.材質, SN.板厚, SN.幅, SN.長さ"
            gSL_Select = gSL_Select & "      ,ISNULL(SN1.入荷数,0) + ISNULL(JK1.実績数,0) - ISNULL(ZS1.使用量,0) - ISNULL(JK2.実績数,0) AS 入荷数"
            gSL_Select = gSL_Select & "      ,ISNULL(SN1.入荷数,0) AS 入荷数入, ISNULL(JK1.実績数,0) AS 実績数入"
            gSL_Select = gSL_Select & "      ,ISNULL(JK2.実績数,0) AS 実績数出, ISNULL(ZS1.使用量,0) AS 使用量出"
            gSL_Select = gSL_Select & " FROM 材料入荷テーブル SN"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(S1.入荷日,1,6) AS 入荷日, S1.材質, S1.板厚, S1.幅, S1.長さ, SUM(S1.入荷数) AS 入荷数"
            gSL_Select = gSL_Select & "                         FROM  材料入荷テーブル S1"
            gSL_Select = gSL_Select & "                         WHERE S1.入荷日 > '" & Trim(ZZM.実棚日) & "'"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(S1.入荷日,1,6), S1.材質, S1.板厚, S1.幅, S1.長さ) SN1"
          ''gSL_Select = gSL_Select & "      --ON  SUBSTRING(SN1.入荷日,1,6) = SUBSTRING(SN.入荷日,1,6)"
          ''gSL_Select = gSL_Select & "      --AND"
            gSL_Select = gSL_Select & "        ON  SN1.材質 = SN.材質 AND SN1.板厚 = SN.板厚 AND SN1.幅 = SN.幅 AND SN1.長さ = SN.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(J1.実績日,1,6) AS 実績日, J1.材質, J1.板厚, J1.幅, J1.長さ, SUM(J1.実績数) AS 実績数"
            gSL_Select = gSL_Select & "                         FROM  材料実績管理テーブル J1"
            gSL_Select = gSL_Select & "                         WHERE J1.実績区分 = 1 AND J1.実績日 > '" & Trim(ZZM.実棚日) & "'"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J1.実績日,1,6), J1.材質, J1.板厚, J1.幅, J1.長さ) JK1"
          ''gSL_Select = gSL_Select & "      --ON  SUBSTRING(JK1.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
          ''gSL_Select = gSL_Select & "      --AND"
            gSL_Select = gSL_Select & "        ON  JK1.材質 = SN.材質 AND JK1.板厚 = SN.板厚 AND JK1.幅 = SN.幅 AND JK1.長さ = SN.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(J2.実績日,1,6) AS 実績日, J2.材質, J2.板厚, J2.幅, J2.長さ, SUM(J2.実績数) AS 実績数"
            gSL_Select = gSL_Select & "                         FROM  材料実績管理テーブル J2"
            gSL_Select = gSL_Select & "                         WHERE J2.実績区分 <> 1 AND J2.入力種類 <> 'A' AND J2.実績日 > '" & Trim(ZZM.実棚日) & "'"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J2.実績日,1,6), J2.材質, J2.板厚, J2.幅, J2.長さ) JK2"
          ''gSL_Select = gSL_Select & "      --ON  SUBSTRING(JK2.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
          ''gSL_Select = gSL_Select & "      --AND"
            gSL_Select = gSL_Select & "        ON  JK2.材質 = SN.材質 AND JK2.板厚 = SN.板厚 AND JK2.幅 = SN.幅 AND JK2.長さ = SN.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(Z1.作業日,1,6) AS 作業日, Z1.材質, Z1.板厚, Z1.幅, Z1.長さ, SUM(Z1.使用量) AS 使用量"
            gSL_Select = gSL_Select & "                         FROM  材料使用量テーブル Z1"
            gSL_Select = gSL_Select & "                         WHERE Z1.作業日   > '" & Trim(ZZM.実棚日) & "'"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(z1.作業日,1,6), Z1.材質, Z1.板厚, Z1.幅, Z1.長さ) ZS1"
          ''gSL_Select = gSL_Select & "      --ON  SUBSTRING(ZS1.作業日,1,6) = SUBSTRING(SN.入荷日,1,6)"
          ''gSL_Select = gSL_Select & "      --AND"
            gSL_Select = gSL_Select & "        ON  ZS1.材質 = SN.材質 AND ZS1.板厚 = SN.板厚 AND ZS1.幅 = SN.幅 AND ZS1.長さ = SN.長さ"
            gSL_Select = gSL_Select & " WHERE SUBSTRING(SN.入荷日,1,6) >= '" & Trim(iH1_YM) & "' AND SN.材料管理番号 = " & Trim(Str(ZZM.材料管理番号))
            gSL_Select = gSL_Select & " GROUP BY SUBSTRING(SN.入荷日,1,6), SN.材質,   SN.板厚,   SN.幅, SN.長さ"
            gSL_Select = gSL_Select & "         ,SN1.入荷数, JK1.実績数  , JK2.実績数,ZS1.使用量"
            gSL_Select = gSL_Select & " ORDER BY SUBSTRING(SN.入荷日,1,6)"
        End If
    End If
    If ZYTRead(gSL_Select) Then
        Do While ZYT_RDSTS
            wZaikoSu = wZaikoSu + ZYT.入荷数
            Call ZYTReadNext
        Loop
    End If
    wZaikoSu = wMasterSu + wZaikoSu
    H1lb_Zaiko.Caption = Format(wZaikoSu, "#,###0")

    GetZaikoSu = True
End Function
'-----------------------------------------------------------------------------------------------------------' 2010.01.15 add end
