VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Begin VB.Form CXU0930 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "売上年月変更処理"
   ClientHeight    =   11235
   ClientLeft      =   2325
   ClientTop       =   2805
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   825
      Left            =   8190
      ScaleHeight     =   765
      ScaleWidth      =   5220
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1560
      Width           =   5280
      Begin CisBtn_60.CisBtn PB_Del 
         Height          =   585
         Left            =   90
         Top             =   90
         Width           =   1545
         _ExtentX        =   2725
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
         PFSize          =   1
         BtnType         =   4
         BtnCaption      =   "実績取消"
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   2565
         Top             =   90
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
         Left            =   3420
         Top             =   90
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
         Left            =   4275
         Top             =   90
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
      Begin CisBtn_60.CisBtn PB_LOOK 
         Height          =   585
         Left            =   1695
         Top             =   90
         Visible         =   0   'False
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
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   5655
      Left            =   1860
      ScaleHeight     =   5595
      ScaleWidth      =   11580
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3240
      Width           =   11640
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   5445
         Left            =   90
         Top             =   60
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   9604
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
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   4170
            Index           =   1
            Left            =   105
            Top             =   1170
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   7355
            BackColor       =   33023
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
            Begin vsFlexLib.vsFlexArray vsGrid1 
               Height          =   3960
               Left            =   510
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   120
               Width           =   10560
               _Version        =   196608
               _ExtentX        =   18627
               _ExtentY        =   6985
               _StockProps     =   228
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ConvInfo        =   1418783674
               FormatString    =   "日付      |便|契約先          |納入先          |受|   件数|変更前年月|処理日時|"
               Rows            =   13
               Cols            =   9
               FixedCols       =   0
               BackColor       =   16777152
               BackColorBkg    =   12632256
               HighLight       =   0
               ScrollBars      =   2
               RowHeightMin    =   300
               SelectionMode   =   1
               BackColorAlternate=   16777152
            End
            Begin Cis3D_v60.CIS3D CIS3D2 
               Height          =   3975
               Index           =   0
               Left            =   90
               Top             =   105
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   7011
               BackColor       =   0
               ForeColor       =   65535
               Caption         =   "処　　理　　履　　歴"
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
               cBoderWidth     =   2
               cFont3D         =   2
               cAlingnment     =   6
               cPositionX      =   30
               cPositionY      =   430
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   990
            Left            =   10560
            Top             =   120
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   1746
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
            Begin CisText_V60.CisText iB1_Ukeir 
               Height          =   360
               Left            =   150
               TabIndex        =   5
               Top             =   450
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
            Height          =   990
            Left            =   6690
            Top             =   120
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   1746
            ForeColor       =   16711680
            Caption         =   "納 入 先"
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
            Begin CisText_V60.CisText iB1_Torcd 
               Height          =   360
               Left            =   90
               TabIndex        =   4
               Top             =   480
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
            Begin Cis3D_v60.CIS3D B1lb_Torcd 
               Height          =   360
               Left            =   1080
               Top             =   480
               Width           =   2640
               _ExtentX        =   4657
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
         Begin Cis3D_v60.CIS3D Back_Bin 
            Height          =   990
            Left            =   90
            Top             =   120
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   1746
            ForeColor       =   128
            Caption         =   "便  "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   3
            cPositionX      =   60
            cPositionY      =   60
            Begin CisYMD_With_Btn.CisYMDwB iB1_Ymd 
               Height          =   345
               Left            =   135
               TabIndex        =   1
               Top             =   480
               Width           =   1950
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
               Object.Height          =   345
               Object.Width           =   1950
               cSize           =   1
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
            Begin CisText_V60.CisText iB1_Bin 
               Height          =   360
               Left            =   2145
               TabIndex        =   2
               Top             =   480
               Width           =   360
               _ExtentX        =   635
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
               cFdAutoFormat   =   1
               cGFormat        =   "##"
               cILength        =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D1 
               Height          =   315
               Left            =   690
               Top             =   60
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               ForeColor       =   128
               Caption         =   "日   付"
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
               cBoderStyle     =   2
               cPositionY      =   10
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   990
            Index           =   3
            Left            =   2760
            Top             =   120
            Width           =   3930
            _ExtentX        =   6932
            _ExtentY        =   1746
            ForeColor       =   128
            Caption         =   "契　　約　　先"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   60
            Begin CisText_V60.CisText iB1_KTorcd 
               Height          =   360
               Left            =   90
               TabIndex        =   3
               Top             =   480
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
            Begin Cis3D_v60.CIS3D B1lb_KTorcd 
               Height          =   360
               Left            =   1080
               Top             =   480
               Width           =   2670
               _ExtentX        =   4710
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
      End
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   915
      Left            =   1830
      ScaleHeight     =   889.899
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   4260
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1530
      Width           =   4320
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   645
         Index           =   1
         Left            =   90
         Top             =   86
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   1138
         ForeColor       =   128
         Caption         =   "売上年月"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   240
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   390
            Left            =   2400
            TabIndex        =   0
            Top             =   135
            Width           =   1440
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
            cAutoEject      =   0   'False
         End
         Begin Cis3D_v60.CIS3D UC_3D3 
            Height          =   255
            Index           =   0
            Left            =   1215
            Top             =   225
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   450
            ForeColor       =   128
            Caption         =   " [ 変更後 ]"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
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
            cBoderStyle     =   2
            cPositionY      =   10
         End
      End
   End
   Begin Cis3D_v60.CIS3D UC_3D11 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 売上年月変更 】"
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
         Left            =   14400
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
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
         Left            =   11850
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
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   705
      Top             =   -120
   End
   Begin VB.Timer Timer1 
      Left            =   270
      Top             =   -150
   End
   Begin VB.TextBox Dummy 
      Height          =   375
      Left            =   -15
      TabIndex        =   7
      Text            =   "Dummy"
      Top             =   -105
      Width           =   345
   End
   Begin VB.Label BG_Logo 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  '透明
      Caption         =   "SIMPR-A"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   390
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "CXU0930"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  売上年月変更
'**       フォームID    :  CXU0930
'**       処理概要      :
'**
'**       作  成  日    :  2005/06/08   By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl

    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim CisVsGrid       As New CisVsGrid3

    Dim mJiseki         As String

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

    '+--------------------------------+
    '+  グリッド初期設定
    '+--------------------------------+
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 350
        .PatanMax = 1
        .SelectCol = 8
        .InitGet = False
        .Init
    End With
    Call GridInit


'   < シングルプロセスチェック >
    If Not SingleProcess_Check Then PB_END_Click

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    If gLogoDisp Then
       BG_Logo(0).Caption = gLogo
       CisFrm.BGLDisp BG_Logo, 10
    End If
    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 1000   ' 時刻表示間隔
    Timer2.Enabled = True    ' 時刻表示開始

    FormAct = Not FormAct

' 売上処理対象ファイルを求める  1=出荷テーブル,1≠検収テーブル
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 50
    Call IRNRead("", 1)
    mJiseki = Mid(IRN.oyian, 1, 1)
    
    If mJiseki <> "1" Then
       Back_Bin.Caption = ""
       iB1_Bin.Visible = False
    End If
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    Cancel = 1
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》       +
'+----------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Key_Acc = False
    If Not FormAct Then Exit Function
    If UCase(IVENT) <> "DOWN" Then Exit Function

    Select Case Key_Code
        Case vbKeyF3:       If PB_Del.Visible Then Call PB_Del_Click        '【実績取消】
        Case vbKeyF7:       If PB_LOOK.Visible Then Call PB_Look_Click      '【検索】
        Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_Can_Click        '【取消】
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click        '【終了】
        Case vbKeyReturn
                            If PB_ENT.Visible Then                          '【入力】
                                Call PB_ENT_Click
                            End If
        Case Else:          Exit Function
    End Select

    Key_Acc = True
End Function
'+----------------------------------------------------------+
'+      フォーム    《Ｐａｉｎｔ イベント》                 +
'+----------------------------------------------------------+
Private Sub Form_Paint()
'ｸﾞﾗﾃﾞｰｼｮﾝ表示
    With CisFrm
        .Gradation_Direction = Portrait
        .Gradation_Shading = Lighten
        .Gradation 250, 100, 100, 2, 5, 10
    End With
End Sub

'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    
    Call SingleProcess_Release
    
    Call CisDB.DBDISConnect
    Call CisFrm.UnLoadDisp
    End
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_Can_Click()
    Call CisFrm.MousePT(11)
   
    If ProcHB = "B1" Then
       Call HeadBodyClear("B1")
       Call CisVsGrid.Clear
       ProcHB = "H1"
       Call DispChange(ProcHB)
       GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)    +
'+----------------------------+
Private Sub PB_ENT_Click()
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call ReturnPress
    Dummy.Enabled = False
End Sub
'+---------------------+
'+    実績取消(F3)     +
'+---------------------+
Private Sub PB_Del_Click()
    CXU0935.Show vbModal
    Unload CXU0935
    Set CXU0935 = Nothing
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
'
    If PB_LOOK.Tag = "iB1_Ymd" Then iB1_Ymd.ShowCalender: Exit Sub
    If PB_LOOK.Tag = "iB1_KTorcd" Then Call Look_KTorcd: Exit Sub
    If PB_LOOK.Tag = "iB1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_LOOK.Tag = "iB1_Ukeir" Then Call Look_Ukeir: Exit Sub

End Sub
'+---------------------+
'+    契約先検索        +
'+---------------------+
Private Sub Look_KTorcd()

    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 2

    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    iB1_KTorcd.SetFocus

    If RV_Rtn Then
       iB1_KTorcd = RV_TorcdK
       Call iB1_KTorcd_LostFocus
    End If

End Sub
'+----------------------------+
'+       納入先検索            +
'+----------------------------+
Private Sub Look_Torcd()

    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 3
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iB1_Torcd = RTrim(RV_TorcdK)
        Call iB1_Torcd_LostFocus
    End If
End Sub
'+----------------------------+
'+     納入先受入検索          +
'+----------------------------+
Private Sub Look_Ukeir()

    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 3
    RV_Torcd = RTrim(iB1_Torcd)
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iB1_Torcd = RTrim(RV_TorcdK)
        Call iB1_Torcd_LostFocus
        iB1_Ukeir = RTrim(RV_Ukeir)
    End If
End Sub
'+-----------------------------+
'+  iB1_Ymd 日付
'+-----------------------------+
Private Sub iB1_Ymd_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iB1_Ymd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_KTorcd   契約先
'+-----------------------------+
Private Sub iB1_KTorcd_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iB1_KTorcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
    B1lb_KTorcd = ""
    If TorNmGet(iB1_KTorcd, 2) Then
        B1lb_KTorcd = TRM.略称
    End If
End Sub
'+-----------------------------+
'+  iB1_Torcd   納入先
'+-----------------------------+
Private Sub iB1_Torcd_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iB1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
    B1lb_Torcd = ""
    If TorNmGet(iB1_Torcd, 3) Then
        B1lb_Torcd = TRM.略称
    End If
End Sub
'+-----------------------------+
'+  iB1_Ukeir
'+-----------------------------+
Private Sub iB1_Ukeir_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iB1_Ukeir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
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
    H1Mode = False: H1Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
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


    PB_CAN.Visible = Not H1Mode
    PB_Del.Visible = H1Mode
    
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
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    
    Exit Sub
HeadBodyClear_ER:
End Sub
'*********************************
'*          入力処理(Enter)      *
'*********************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        
        Call GridSet
        
        ProcHB = "B1"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B1")
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

    If Trim(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　売上年月[変更後]を入力して下さい           "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If
    
    If Not iH1_YM.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　正しい年月にて入力して下さい           "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If
    Head1Chk = True
End Function
'****************************
'*      BODY1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean

    Body1Chk = False

    If Trim(iB1_Ymd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　日付を入力して下さい           "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_Ymd.SetFocus
        Exit Function
    End If
    If Not iB1_Ymd.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　正しい日付にて入力して下さい           "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_Ymd.SetFocus
        Exit Function
    End If
    If mJiseki = "1" Then
        If CisFun.Val2(iB1_Bin) = 0 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "   　便を入力して下さい           "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB1_Bin.SetFocus
            Exit Function
        End If
    End If
    If Trim(iB1_KTorcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　契約先を入力して下さい           "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_KTorcd.SetFocus
        Exit Function
    End If

    If Not TorNmGet(iB1_KTorcd, 2) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　契約先　未登録    "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_KTorcd.SetFocus
        Exit Function
    End If
    
    If Trim(iB1_Torcd) <> "" Then
        If Not TorNmGet(iB1_Torcd, 3) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "   　納入先　未登録    "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB1_Torcd.SetFocus
            Exit Function
        End If
    End If

    If Trim(iB1_Ukeir) <> "" And Trim(iB1_Torcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　納入先を入力して下さい           "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_Torcd.SetFocus
        Exit Function
    End If
    
    Body1Chk = True
End Function
'
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
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
        Call .FixedSet(中中, 中中, あり, 10, 0, "日付")
        Call .FixedSet(中中, 中中, あり, 2, 1, "便")
        Call .FixedSet(中中, 左中, あり, 16, 2, "契約先")
        Call .FixedSet(中中, 左中, あり, 16, 3, "納入先")
        Call .FixedSet(中中, 中中, あり, 2, 4, "受")
        If mJiseki <> "1" Then
           Call .FixedSet(中中, 右中, あり, 6, 5, "件数")
        Else
           Call .FixedSet(中中, 右中, あり, 8, 5, "件数")
        End If
        Call .FixedSet(中中, 中中, あり, 0, 6, "変更前年月")
        Call .FixedSet(中中, 中中, あり, 18, 7, "処理日時")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 8, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        If mJiseki <> "1" Then
           Call .PatanSet(0, "7,0,1,2,3,4,5,6")
        Else
           Call .PatanSet(0, "7,0,2,3,4,5,6")
        End If
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

    Call CisVsGrid.Clear
' 抽出処理 初期ｾｯﾄ

    gSL_Select = "Select UH.*, "
    gSL_Select = gSL_Select & "SUBSTRING(CONVERT(CHAR(10),UH.作成日,111),3,8) + ' ' + CONVERT(CHAR(8),UH.作成日,114) 処理日時,"
    gSL_Select = gSL_Select & " ISNULL(TR.略称,'') 契約先名,"
    gSL_Select = gSL_Select & " ISNULL(TR1.略称,'') 納入先名"
    gSL_Select = gSL_Select & " From 売上変更履歴テーブル UH"
    gSL_Select = gSL_Select & "   Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & "    ON TR.取引先区分 = 0"
    gSL_Select = gSL_Select & "    And TR.取引先CD = UH.契約先"
    gSL_Select = gSL_Select & "   Left Outer Join 取引先マスタ TR1"
    gSL_Select = gSL_Select & "    ON TR1.取引先区分 = 0"
    gSL_Select = gSL_Select & "    And TR1.取引先CD = UH.納入先"
    gSL_Select = gSL_Select & " Where UH.売上年月 = '" & iH1_YM & "'"
    gSL_Select = gSL_Select & "   And UH.処理区分 = ''"
    gSL_Select = gSL_Select & "   And UH.削除区分 = ''"
    gSL_Select = gSL_Select & " Order By UH.作成日 Desc,UH.日付,UH.便,UH.契約先,UH.納入先,UH.受入"
    
    If Not UHRRead(gSL_Select) Then
       Call UHRClose
       Exit Sub
    End If

    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
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
        ID = 0
        Do Until Not UHR_RDSTS
            ID = ID + 1
            '日付
            With CisFun
                .DateE_BefVal = UHR.日付
                .DateEdit
                vsGrid1.TextMatrix(ID, GridCol(0)) = .DateE_AfVal
            End With
            '便
            .TextMatrix(ID, GridCol(1)) = UHR.便
            '契約先
            gStr = UHR.契約先
            If Trim(UHR.契約先名) <> "" Then
               gStr = gStr & ":" & UHR.契約先名
            End If
            .TextMatrix(ID, GridCol(2)) = gStr
            '納入先
            gStr = UHR.納入先
            If Trim(UHR.納入先名) <> "" Then
               gStr = gStr & ":" & UHR.納入先名
            End If
            .TextMatrix(ID, GridCol(3)) = gStr
            '受入
            .TextMatrix(ID, GridCol(4)) = UHR.受入
            '件数
            .TextMatrix(ID, GridCol(5)) = Format(UHR.処理件数, "#,##0")
            '変更前年月
            If Trim(UHR.変更前年月) <> "" Then
               .TextMatrix(ID, GridCol(6)) = Mid(UHR.変更前年月, 1, 4) & "/" & Mid(UHR.変更前年月, 5, 2)
            End If
            
            '処理日
            .TextMatrix(ID, GridCol(7)) = UHR.処理日時
            
            Call UHRReadNext
        Loop
        Call UHRClose
        
        .Row = 1: .Col = 1: .ColSel = .Cols - 1
        .Redraw = True
    End With
End Sub
'**************************************
'*                                    *
'*          データベース更新          *
'*                                    *
'**************************************
'+----------------------------+
'+      削 除 処 理            +
'+----------------------------+
Private Function DBDelete()
    DBDelete = False
''  ( 確認メッセージ )
'    K_Sykbnm.Caption = SyoriNM
'    K_Sykbnm.Visible = True
'    With CisFun
'        .MB_Lines = 4
'        .MB_MSG(2) = "     削除処理を行います      "
'        .MB_Title = "削除処理"
'        .MB_MSG(3) = "     よろしいですか？      "
'        .MB_Button = OK_CAN
'        If Not .MBOX Then Exit Function
'    End With
'
'    With CisDB
'        On Error GoTo DBDelete_Err
'        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
'' ( 削除処理 )
'        .SQL = "DELETE FROM 売上実績テーブル "
'        .SQL = .SQL & " WHERE 検収管理NO = " & CisFun.Val2(vsGrid1.TextMatrix(vsGrid1.Row, CisVsGrid1.FixedGet(14))) & ""
'        .DBExec
'
'        If B1Chk_Syogo.Value = 1 Then
'            .SQL = "DELETE FROM 売上明細テーブル "
'            .SQL = .SQL & " WHERE 検収管理NO = " & CisFun.Val2(vsGrid1.TextMatrix(vsGrid1.Row, CisVsGrid1.FixedGet(14))) & ""
'            .DBExec
'        End If
'
'        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
'
'        iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")
'
'        On Error GoTo 0
'    End With
'
''Grid再表示
'    Call vsGrid1_ReSetDisp
    
    DBDelete = True
    Exit Function
DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加･更新 処理    +
'+----------------------+
Private Function DBPut() As Boolean
    DBPut = False
    Dim wWhere  As String
    
    On Error GoTo DBPut_Err

'
    If mJiseki <> "1" Then
        gSL_Select = "SELECT KN.検収日,HT.契約先"
        If Trim(iB1_Torcd) <> "" Then
           gSL_Select = gSL_Select & "   ,KN.納入先"
        End If
        If Trim(iB1_Ukeir) <> "" Then
           gSL_Select = gSL_Select & "   ,KN.受入"
        End If
        gSL_Select = gSL_Select & ",Count(*) 件数,MAX(KN.実績年月) 実績年月"
        gSL_Select = gSL_Select & "  From 検収テーブル KN"
        gSL_Select = gSL_Select & "    Left Outer Join 品番取引先マスタ HT"
        gSL_Select = gSL_Select & "     ON  HT.品番 = KN.品番"
        gSL_Select = gSL_Select & "     And HT.取引先 = KN.納入先"
        gSL_Select = gSL_Select & "     And HT.受入  = KN.受入"
        ' 抽出条件
        wWhere = " Where KN.検収日 = '" & iB1_Ymd & "'"
        wWhere = wWhere & "   And HT.契約先 = '" & iB1_KTorcd & "'"
        If Trim(iB1_Torcd) <> "" Then
           wWhere = wWhere & "   And KN.納入先 = '" & iB1_Torcd & "'"
        End If
        If Trim(iB1_Ukeir) <> "" Then
           wWhere = wWhere & "   And KN.受入 = '" & iB1_Ukeir & "'"
        End If
        
        gSL_Select = gSL_Select & wWhere
        
        gSL_Select = gSL_Select & " Group By KN.検収日,HT.契約先"
        If Trim(iB1_Torcd) <> "" Then
           gSL_Select = gSL_Select & "   ,KN.納入先"
        End If
        If Trim(iB1_Ukeir) <> "" Then
           gSL_Select = gSL_Select & "   ,KN.受入"
        End If
    Else
        gSL_Select = "SELECT JS.出荷日 検収日,HT.契約先"
        If Trim(iB1_Torcd) <> "" Then
           gSL_Select = gSL_Select & "   ,KN.納入先"
        End If
        If Trim(iB1_Ukeir) <> "" Then
           gSL_Select = gSL_Select & "   ,KN.受入"
        End If
        gSL_Select = gSL_Select & ",Count(*) 件数,MAX(JS.実績年月) 実績年月"
        gSL_Select = gSL_Select & "   From 出荷テーブル JS"
        gSL_Select = gSL_Select & "     LEFT OUTER JOIN 受注テーブル KN"
        gSL_Select = gSL_Select & "        ON  KN.出荷指示書NO = JS.出荷指示書NO"
        gSL_Select = gSL_Select & "        AND KN.出荷指示書行NO = JS.出荷指示書行NO"
        gSL_Select = gSL_Select & "      Left Outer Join 品番取引先マスタ HT"
        gSL_Select = gSL_Select & "        ON  HT.品番 = KN.品番"
        gSL_Select = gSL_Select & "        And HT.取引先 = KN.納入先"
        gSL_Select = gSL_Select & "        And HT.受入  = KN.受入"
        
        ' 抽出条件
        wWhere = " Where JS.出荷日 = '" & iB1_Ymd & "'"
        wWhere = wWhere & "   And JS.出荷便 = '" & iB1_Bin & "'"
        wWhere = wWhere & "   And HT.契約先 = '" & iB1_KTorcd & "'"
        If Trim(iB1_Torcd) <> "" Then
           wWhere = wWhere & "   And KN.納入先 = '" & iB1_Torcd & "'"
        End If
        If Trim(iB1_Ukeir) <> "" Then
           wWhere = wWhere & "   And KN.受入 = '" & iB1_Ukeir & "'"
        End If
        
        gSL_Select = gSL_Select & wWhere
        
        gSL_Select = gSL_Select & " Group By JS.出荷日,JS.出荷便,HT.契約先"
        If Trim(iB1_Torcd) <> "" Then
           gSL_Select = gSL_Select & "   ,KN.納入先"
        End If
        If Trim(iB1_Ukeir) <> "" Then
           gSL_Select = gSL_Select & "   ,KN.受入"
        End If
    End If

    If Not KNTRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　   対象データがありません。             "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If
' 実績変更確認
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     実績年月変更処理を行います      "
        .MB_Title = "更新処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With

    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

' ( 実績取消処理 )
    With CisDB
        If mJiseki <> "1" Then
           wWhere = Replace(wWhere, "KN.", "検収テーブル.")
           wWhere = Replace(wWhere, "Where", "And")
        
           .SQL = "Update 検収テーブル Set "
           .SQL = .SQL & " 実績年月 = '" & iH1_YM & "',"
           .SQL = .SQL & " 更新日 = GETDATE(),"
           .SQL = .SQL & " 更新端末 = SubString(Host_Name(),1,20),"
           .SQL = .SQL & " 更新者 = '" & RTrim(gTanto) & "'"
           .SQL = .SQL & "   From 品番取引先マスタ HT"
           .SQL = .SQL & "     Where HT.品番 = 検収テーブル.品番"
           .SQL = .SQL & "       And HT.取引先 = 検収テーブル.納入先"
           .SQL = .SQL & "       And HT.受入  = 検収テーブル.受入"
           .SQL = .SQL & wWhere
        Else
           wWhere = Replace(wWhere, "JS.", "出荷テーブル.")
           wWhere = Replace(wWhere, "Where", "And")
        
           .SQL = "Update 出荷テーブル Set "
           .SQL = .SQL & " 実績年月 = '" & iH1_YM & "',"
           .SQL = .SQL & " 更新日 = GETDATE(),"
           .SQL = .SQL & " 更新端末 = SubString(Host_Name(),1,20),"
           .SQL = .SQL & " 更新者 = '" & RTrim(gTanto) & "'"
           .SQL = .SQL & "   From 受注テーブル KN"
           .SQL = .SQL & "   Left Outer Join 品番取引先マスタ HT"
           .SQL = .SQL & "       On  HT.品番 = KN.品番"
           .SQL = .SQL & "       And HT.取引先 = KN.納入先"
           .SQL = .SQL & "       And HT.受入  = KN.受入"
           .SQL = .SQL & " Where  KN.出荷指示書NO = 出荷テーブル.出荷指示書NO"
           .SQL = .SQL & "   AND KN.出荷指示書行NO = 出荷テーブル.出荷指示書行NO"
           .SQL = .SQL & wWhere
        End If
        If Not .DBExec Then GoTo DBPut_Err
    End With
' 変更履歴出力
    Call ItemsClearUHR
    With UHR
        .売上年月 = iH1_YM
        .日付 = iB1_Ymd
        .便 = iB1_Bin
        .契約先 = iB1_KTorcd
        .納入先 = iB1_Torcd
        .受入 = iB1_Ukeir
        .処理件数 = KNT.件数
        .変更前年月 = KNT.実績年月
        .作成元区分 = mJiseki
        .作成者 = gTanto
    End With
    Call UHRInsert
'
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

''Grid再表示
    Call GridSet
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'*---------------------------------------------------------*
'*          シ ン グ ル プ ロ セ ス チ ェ ッ ク             *
'*---------------------------------------------------------*
Private Function SingleProcess_Check()
    SingleProcess_Check = False
    
    Call ItemsClearSYMCreate
    
    With SYMCreate
        .区分名称 = "シングルプロセス"
        .区分桁数 = 10
        .値 = App.EXEName
        .制御区分 = "X"
        .値名称 = "売上変更処理"
        .作成者 = gTanto
    End With
    
    Call SYMInitCreate(False)
    
    Call ItemsClearSYM
    gSL_Select = "Select 英数字1 From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称   = 'シングルプロセス'"
    gSL_Select = gSL_Select & " And   区分タイプ = ''"
    gSL_Select = gSL_Select & " And   値        = '" & App.EXEName & "'"
    Call SYMRead(gSL_Select, 1)
    
    With CisFun
        If RTrim$(SYM.英数字1) <> "" And RTrim$(SYM.英数字1) <> .WSID Then
            .MB_Lines = 3
            .MB_MSG(1) = "    他のパソコンで起動中です         "
            .MB_MSG(2) = "        ( " & RTrim$(SYM.英数字1) & " )"
            .MB_Button = Error
            .MBOX
            Exit Function
        End If
    End With

'   < プロセスのロック >
    gSL_Select = "UpDate 名称マスタ Set"
    gSL_Select = gSL_Select & " 英数字1 = SubString( Host_Name(), 1, 20 ),"
    gSL_Select = gSL_Select & " 更新日 = GETDATE(),"
    gSL_Select = gSL_Select & " 更新者 = '" & RTrim(gTanto) & "',"
    gSL_Select = gSL_Select & " 更新端末 = SubString( Host_Name(), 1, 20 )"
    gSL_Select = gSL_Select & " Where 区分名称 = 'シングルプロセス'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   And 値 = '" & App.EXEName & "'"

    With CisDB
        .SQL = gSL_Select
        .DBExec
    End With

    SingleProcess_Check = True
End Function
'*-------------------------------------------------*
'*          シ ン グ ル プ ロ セ ス 解 除           *
'*-------------------------------------------------*
Private Function SingleProcess_Release()
    SingleProcess_Release = False
    
    gSL_Select = "UpDate 名称マスタ Set"
    gSL_Select = gSL_Select & " 英数字1 = '',"
    gSL_Select = gSL_Select & " 更新日 = GETDATE(),"
    gSL_Select = gSL_Select & " 更新者 = '" & RTrim(gTanto) & "',"
    gSL_Select = gSL_Select & " 更新端末 = SubString( Host_Name(), 1, 20 )"
    gSL_Select = gSL_Select & " Where 区分名称 = 'シングルプロセス'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   And 値 = '" & App.EXEName & "'"
    gSL_Select = gSL_Select & "   And 英数字1 = SubString( Host_Name(), 1, 20 )"

    With CisDB
        .SQL = gSL_Select
        .DBExec
    End With
    
    SingleProcess_Release = True
End Function



